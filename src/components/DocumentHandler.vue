<template>
    <div class="editor-container" v-loading="loading" element-loading-text="Loading...">
        <div id="iframe"></div>
    </div>
</template>

<script lang="ts" setup>
import { onMounted, onBeforeUnmount, ref, watch } from 'vue'
import { getDocumentType, DocmentType } from '@/utils/util'
import { g_sEmpty_bin } from '@/utils/empty_bin'
// @ts-ignore
import {
    initX2TScript,
    initX2T,
    convertDocument,
    convertBinToDocument,
    c_oAscFileType2,
    oAscFileType,
} from '@/utils/x2t'
const X2T = ref(null)
// 设置prop
const props = defineProps<{
    file: DocmentType
    /** 只读预览：禁止编辑与保存 */
    readonly?: boolean
}>()

const emit = defineEmits<{
    onlyofficeCursor: [payload: { event: string; source?: string; doc?: Record<string, unknown> }]
}>()

const CURSOR_PLUGIN_GUID = 'asc.{6C757CD6-7A91-4B0F-9C2D-1E3F4A5B6C7D}'

function getCursorPluginConfigUrl(): string {
    try {
        let base = import.meta.env.BASE_URL || '/'
        if (!base.endsWith('/')) base += '/'
        return new URL('onlyoffice-cursor-plugin/config.json', new URL(base, window.location.origin)).href
    } catch {
        return `${window.location.origin}/onlyoffice-cursor-plugin/config.json`
    }
}

/** 本地 web-apps 常未真正加载自定义插件，改为在编辑器 iframe 内订阅 Asc.editor 事件（同源） */
let innerBridgeTimer: ReturnType<typeof setInterval> | null = null
let innerBridgeCleanup: (() => void) | null = null

function clearInnerCursorBridge() {
    if (innerBridgeTimer) {
        clearInterval(innerBridgeTimer)
        innerBridgeTimer = null
    }
    if (innerBridgeCleanup) {
        innerBridgeCleanup()
        innerBridgeCleanup = null
    }
}

/** 与任意父页面约定：type === ONLYOFFICE_CURSOR_PROBE，payload 含 event / source / doc */
function postCursorProbe(event: string, source: string, api?: Record<string, unknown>) {
    const doc = api ? collectEditorCursorSnapshot(api) : undefined
    const msg = { type: 'ONLYOFFICE_CURSOR_PROBE', payload: { event, source, doc } }
    window.postMessage(msg, '*')
    try {
        if (window.top && window.top !== window) {
            window.top.postMessage(msg, '*')
        }
    } catch {
        /* 跨域 top 不可写时忽略 */
    }
}

function postDocumentModified(modified = true) {
    if (!modified) return
    const msg = { type: 'ONLYOFFICE_DOCUMENT_MODIFIED', payload: { modified: true } }
    window.postMessage(msg, '*')
    try {
        if (window.parent && window.parent !== window) {
            window.parent.postMessage(msg, '*')
        }
    } catch {
        /* ignore */
    }
}

/** 使用编辑器内置 pluginMethod（与插件同源），读取选区文本 / 类型；并读取插入点/选区在版式中的坐标（非浏览器像素） */
function serializeAnchorRect(pos: unknown): Record<string, number> | null {
    if (!pos || typeof pos !== 'object') return null
    const p = pos as Record<string, unknown>
    const ax = p.asc_getX
    const ay = p.asc_getY
    const aw = p.asc_getWidth
    const ah = p.asc_getHeight
    if (
        typeof ax === 'function' &&
        typeof ay === 'function' &&
        typeof aw === 'function' &&
        typeof ah === 'function'
    ) {
        return {
            x: (ax as (this: unknown) => number).call(pos),
            y: (ay as (this: unknown) => number).call(pos),
            width: (aw as (this: unknown) => number).call(pos),
            height: (ah as (this: unknown) => number).call(pos),
        }
    }
    return null
}

function collectEditorCursorSnapshot(api: Record<string, unknown>): Record<string, unknown> {
    const out: Record<string, unknown> = {}
    try {
        const getSel = api.pluginMethod_GetSelectedText
        const getType = api.pluginMethod_GetSelectionType
        if (typeof getSel === 'function') {
            const selectedText = (getSel as (this: unknown, opts?: object) => unknown).call(api, {
                Numbering: true,
                Math: true,
            })
            out.selectedText = selectedText
            out.selectedLength =
                typeof selectedText === 'string'
                    ? selectedText.length
                    : Array.isArray(selectedText)
                      ? JSON.stringify(selectedText).length
                      : 0
        }
        if (typeof getType === 'function') {
            out.selectionType = (getType as (this: unknown) => unknown).call(api)
        }
        if (api.cn !== undefined) out.editorCn = api.cn

        const len = typeof out.selectedLength === 'number' ? out.selectedLength : 0
        out.isCollapsed = len === 0

        const getAnchor = api.asc_getAnchorPosition
        if (typeof getAnchor === 'function') {
            try {
                const pos = (getAnchor as (this: unknown) => unknown).call(api)
                const rect = serializeAnchorRect(pos)
                if (rect) out.anchor = rect
            } catch (e) {
                out.anchorError = String(e)
            }
        }

        const getBounds = api.asc_GetSelectionBounds
        if (typeof getBounds === 'function') {
            try {
                out.selectionBounds = (getBounds as (this: unknown) => unknown).call(api)
            } catch (e) {
                out.selectionBoundsError = String(e)
            }
        }
    } catch (e) {
        out.snapshotError = String(e)
    }
    return out
}

type EditorIframeWindow = Window & {
    Asc?: { editor?: Record<string, unknown>; aCa?: new (fileType: number, isDownload?: boolean) => object }
    editor?: Record<string, unknown>
}

function getEditorIframeWindow(): EditorIframeWindow | null {
    const iframe = document.querySelector('.editor-container iframe') as HTMLIFrameElement | null
    return (iframe?.contentWindow as EditorIframeWindow | undefined) ?? null
}

function getEditorApiFromDom(): Record<string, unknown> | null {
    const w = getEditorIframeWindow()
    const api = (w?.Asc?.editor ?? w?.editor) as Record<string, unknown> | undefined
    return api ?? null
}

/** 在当前光标处插入纯文本（Word/表格等内置支持 pluginMethod_InputText 时可用） */
function insertTextAtCursor(text: string): boolean {
    const api = getEditorApiFromDom()
    if (!api || typeof api.pluginMethod_InputText !== 'function') {
        console.warn('[OnlyOffice] pluginMethod_InputText 不可用（可能未打开文档或非文本模式）')
        return false
    }
    try {
        ;(api.pluginMethod_InputText as (this: unknown, t: string, replace: string) => void).call(
            api,
            String(text ?? ''),
            '',
        )
        return true
    } catch (e) {
        console.error('[OnlyOffice] insertTextAtCursor 失败', e)
        return false
    }
}

/** 在当前光标处粘贴 HTML 片段（需 Word 可编辑且 pluginMethod_PasteHtml 可用） */
function insertHtmlAtCursor(html: string): boolean {
    const api = getEditorApiFromDom()
    if (!api || typeof api.pluginMethod_PasteHtml !== 'function') {
        console.warn('[OnlyOffice] pluginMethod_PasteHtml 不可用')
        return false
    }
    try {
        ;(api.pluginMethod_PasteHtml as (this: unknown, h: string) => void).call(api, String(html ?? ''))
        return true
    } catch (e) {
        console.error('[OnlyOffice] insertHtmlAtCursor 失败', e)
        return false
    }
}

/** 整篇 HTML：编辑器 iframe 内 asc_nativeGetHtml（当前环境无 ToHtml 时用此路径） */
function getDocumentHtmlSync(): string | null {
    const api = getEditorApiFromDom()
    if (!api) return null
    if (typeof api.asc_nativeGetHtml === 'function') {
        try {
            const html = (api.asc_nativeGetHtml as (this: unknown) => string).call(api)
            if (typeof html === 'string' && html.trim()) {
                console.info('[OnlyOffice] getDocumentHtml: asc_nativeGetHtml', html.length)
                return html
            }
        } catch (e) {
            console.warn('[OnlyOffice] asc_nativeGetHtml', e)
        }
    }
    if (typeof api.pluginMethod_GetFileHTML === 'function') {
        try {
            const html = (api.pluginMethod_GetFileHTML as (this: unknown) => unknown).call(api)
            if (typeof html === 'string' && html.trim()) {
                console.info('[OnlyOffice] getDocumentHtml: pluginMethod_GetFileHTML', html.length)
                return html
            }
        } catch (e) {
            console.warn('[OnlyOffice] pluginMethod_GetFileHTML', e)
        }
    }
    console.warn('[OnlyOffice] getDocumentHtml: asc_nativeGetHtml / GetFileHTML 均不可用')
    return null
}

function getDocumentHtml(callback: (html: string | null) => void) {
    callback(getDocumentHtmlSync())
}

/** 在文档中查找并选中下一处匹配（对齐 web-apps 查找框：CSearchSettings + asc_findText） */
function findTextInDocument(text: string, forward = true): boolean {
    const query = String(text ?? '').trim()
    if (!query) return false

    const api = getEditorApiFromDom()
    const w = getEditorIframeWindow()
    if (!api || typeof api.asc_findText !== 'function') {
        console.warn('[OnlyOffice] asc_findText 不可用')
        return false
    }

    const CSearchSettings = (w as EditorIframeWindow & { AscCommon?: { CSearchSettings?: new () => SearchSettingsLike } })
        ?.AscCommon?.CSearchSettings
    if (typeof CSearchSettings !== 'function') {
        console.warn('[OnlyOffice] AscCommon.CSearchSettings 不可用')
        return false
    }

    try {
        const settings = new CSearchSettings()
        settings.put_Text(query)
        settings.put_MatchCase(false)
        settings.put_WholeWords(false)
        const found = !!(api.asc_findText as (s: unknown, fwd: boolean) => boolean).call(api, settings, forward)
        if (found && typeof api.asc_StartTextAroundSearch === 'function') {
            try {
                ;(api.asc_StartTextAroundSearch as (this: unknown) => void).call(api)
            } catch {
                /* ignore */
            }
        }
        return found
    } catch (e) {
        console.warn('[OnlyOffice] findTextInDocument', e)
        return false
    }
}

type SearchSettingsLike = {
    put_Text: (t: string) => void
    put_MatchCase: (v: boolean) => void
    put_WholeWords: (v: boolean) => void
}

defineExpose({
    insertTextAtCursor,
    insertHtmlAtCursor,
    getDocumentHtml,
    findTextInDocument,
    getEditorApiFromDom,
})

/** 父页 REQUEST_SAVE 等待 onSave 完成后回传 docx */
const pendingSaveReplies = new Map<
    string,
    (payload: { fileName: string; bytes: Uint8Array } | null, err?: string) => void
>()

function postSaveDocumentReply(
    target: MessageEventSource | null,
    requestId: string | undefined,
    payload: { fileName: string; bytes: Uint8Array } | null,
    error?: string,
) {
    const reply = {
        type: 'ONLYOFFICE_SAVE_DOCUMENT_REPLY',
        payload: { requestId, fileName: payload?.fileName, bytes: payload?.bytes, error },
    }
    try {
        ;(target as Window | null)?.postMessage?.(reply, '*')
    } catch {
        /* ignore */
    }
    try {
        if (window.parent && window.parent !== window) {
            window.parent.postMessage(reply, '*')
        }
    } catch {
        /* ignore */
    }
}

function flushPendingSaveReply(
    requestId: string | undefined,
    payload: { fileName: string; bytes: Uint8Array } | null,
    error?: string,
) {
    if (requestId) {
        const fn = pendingSaveReplies.get(requestId)
        if (fn) {
            pendingSaveReplies.delete(requestId)
            fn(payload, error)
            return
        }
    }
    for (const [id, fn] of pendingSaveReplies.entries()) {
        pendingSaveReplies.delete(id)
        fn(payload, error)
        break
    }
}

type SavePayload = { fileName: string; bytes: Uint8Array }

async function blobToSavePayload(blob: Blob, fileName: string): Promise<SavePayload> {
    const buf = await blob.arrayBuffer()
    return { fileName, bytes: new Uint8Array(buf) }
}

async function resolveDownloadDataToPayload(data: string, fileName: string): Promise<SavePayload> {
    if (data.startsWith('http://') || data.startsWith('https://') || data.startsWith('blob:')) {
        const res = await fetch(data)
        if (!res.ok) throw new Error(`拉取导出文件失败: ${res.status}`)
        return blobToSavePayload(await res.blob(), fileName)
    }
    const binary = atob(data.replace(/\s/g, ''))
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
    return { fileName, bytes }
}

function finishAhaDownload(
    result: unknown,
    fileName: string,
    resolve: (p: SavePayload) => void,
    reject: (e: Error) => void,
) {
    const r = result as { status?: string; data?: string } | null
    if (r?.status === 'ok' && r?.data && r.data !== 'error') {
        void resolveDownloadDataToPayload(String(r.data), fileName).then(resolve).catch(reject)
        return
    }
    reject(new Error('编辑器导出失败'))
}

/** SDK 在 pluginMethod_GetFileToDownload / asc_DownloadAs 内部异步设置 Aha，须在调用后轮询包装 */
function exportWithApiAha(
    api: Record<string, unknown>,
    fileName: string,
    invoke: () => void,
    timeoutMs: number,
): Promise<SavePayload> {
    return new Promise((resolve, reject) => {
        let wrapped = false
        const cleanup = () => {
            clearTimeout(timer)
            clearInterval(poll)
        }
        const timer = setTimeout(() => {
            cleanup()
            reject(new Error('导出超时'))
        }, timeoutMs)
        const poll = setInterval(() => {
            const inner = api.Aha
            if (wrapped || typeof inner !== 'function') return
            wrapped = true
            api.Aha = (result: unknown) => {
                cleanup()
                try {
                    ;(inner as (this: unknown, r: unknown) => void).call(api, result)
                } catch {
                    /* ignore */
                }
                finishAhaDownload(result, fileName, resolve, reject)
            }
        }, 16)
        try {
            invoke()
        } catch (e) {
            cleanup()
            reject(e instanceof Error ? e : new Error(String(e)))
        }
    })
}

function dispatchEditorSaveShortcut(): boolean {
    const iframe = document.querySelector('.editor-container iframe') as HTMLIFrameElement | null
    const doc = iframe?.contentWindow?.document
    if (!doc) return false
    const opts = { key: 's', code: 'KeyS', ctrlKey: true, bubbles: true, cancelable: true }
    doc.dispatchEvent(new KeyboardEvent('keydown', opts))
    doc.dispatchEvent(new KeyboardEvent('keyup', opts))
    return true
}

function triggerEditorSaveCommands() {
    const api = getEditorApiFromDom()
    if (api && typeof api.asc_Save === 'function') {
        try {
            ;(api.asc_Save as (this: unknown) => void).call(api)
        } catch (e) {
            console.warn('[OnlyOffice] asc_Save failed', e)
        }
    }
    if (editor.value && typeof editor.value.sendCommand === 'function') {
        editor.value.sendCommand({ command: 'asc_Save', data: {} })
    }
    dispatchEditorSaveShortcut()
}

function rawEditorBinToUint8(raw: unknown): Uint8Array | null {
    if (raw instanceof Uint8Array) return raw.length ? raw : null
    if (raw instanceof ArrayBuffer) return raw.byteLength ? new Uint8Array(raw) : null
    if (typeof raw === 'string' && raw.length > 0) {
        try {
            const bin = Uint8Array.from(atob(raw.replace(/\s/g, '')), (c) => c.charCodeAt(0))
            return bin.length ? bin : null
        } catch {
            return null
        }
    }
    if (raw && typeof raw === 'object' && 'data' in raw) {
        return rawEditorBinToUint8((raw as { data: unknown }).data)
    }
    return null
}

/** 浏览器内从内存取 bin（勿用 asc_nativeGetFileData，其依赖 native.Save_End） */
async function exportViaMemoryBin(fileName: string): Promise<SavePayload | null> {
    const api = getEditorApiFromDom()
    if (!api) return null
    const getters = [api.asc_nativeGetFile2, api.asc_nativeGetFile].filter(
        (fn): fn is (this: unknown) => unknown => typeof fn === 'function',
    )
    for (const getBin of getters) {
        try {
            const raw = (getBin as (this: unknown) => unknown).call(api)
            const bin = rawEditorBinToUint8(raw)
            if (!bin) continue
            const converted = await convertBinToDocument(bin, fileName, 'DOCX')
            return { fileName: converted.fileName, bytes: converted.data }
        } catch (e) {
            console.warn('[OnlyOffice] exportViaMemoryBin', e)
        }
    }
    return null
}

function createDocxDownloadOptions(): object | null {
    const w = getEditorIframeWindow()
    if (!w?.Asc?.aCa) return null
    return new w.Asc.aCa(oAscFileType.DOCX)
}

function exportViaEditorOnSave(fileName: string, timeoutMs: number): Promise<SavePayload> {
    return new Promise((resolve, reject) => {
        if (!editor.value || typeof editor.value.sendCommand !== 'function') {
            reject(new Error('编辑器未就绪'))
            return
        }
        const requestId = `save-${Date.now()}`
        const timer = setTimeout(() => {
            pendingSaveReplies.delete(requestId)
            reject(new Error('onSave 未触发'))
        }, timeoutMs)
        pendingSaveReplies.set(requestId, (payload, err) => {
            clearTimeout(timer)
            if (err || !payload) reject(new Error(err || '保存数据为空'))
            else resolve(payload)
        })
        try {
            editor.value.sendCommand({
                command: 'asc_DownloadAs',
                data: { format: oAscFileType.DOCX, fileType: oAscFileType.DOCX },
            })
        } catch (e) {
            clearTimeout(timer)
            pendingSaveReplies.delete(requestId)
            reject(e instanceof Error ? e : new Error(String(e)))
        }
    })
}

/** 供父页自动保存：多路径导出，避免 asc_Save 不触发 onSave */
async function exportDocumentForSave(fileName: string): Promise<SavePayload> {
    const memory = await exportViaMemoryBin(fileName)
    if (memory) return memory

    const api = getEditorApiFromDom()
    const ext = (fileName.split('.').pop() || 'docx').toUpperCase()

    if (api && typeof api.pluginMethod_GetFileToDownload === 'function') {
        try {
            return await exportWithApiAha(
                api,
                fileName,
                () => {
                    ;(api.pluginMethod_GetFileToDownload as (this: unknown, e: string) => void).call(
                        api,
                        ext,
                    )
                },
                45_000,
            )
        } catch (e) {
            console.warn('[OnlyOffice] pluginMethod_GetFileToDownload', e)
        }
    }

    if (api && typeof api.asc_DownloadAs === 'function') {
        const opts = createDocxDownloadOptions()
        if (opts) {
            try {
                return await exportWithApiAha(
                    api,
                    fileName,
                    () => {
                        ;(api.asc_DownloadAs as (this: unknown, o: object) => void).call(api, opts)
                    },
                    45_000,
                )
            } catch (e) {
                console.warn('[OnlyOffice] asc_DownloadAs', e)
            }
        }
    }

    try {
        return await exportViaEditorOnSave(fileName, 25_000)
    } catch (e) {
        console.warn('[OnlyOffice] exportViaEditorOnSave', e)
    }

    return new Promise((resolve, reject) => {
        const requestId = `save-${Date.now()}`
        const timer = setTimeout(() => {
            pendingSaveReplies.delete(requestId)
            reject(new Error('onSave 未触发'))
        }, 20_000)
        pendingSaveReplies.set(requestId, (payload, err) => {
            clearTimeout(timer)
            if (err || !payload) reject(new Error(err || '保存数据为空'))
            else resolve(payload)
        })
        triggerEditorSaveCommands()
    })
}

/** 同源父页用 ref；最外层 iframe 向本窗口 postMessage：INSERT_* / GET_DOCUMENT_HTML（回复 ONLYOFFICE_GET_DOCUMENT_HTML_REPLY） */
function onCursorPluginMessage(ev: MessageEvent) {
    const d = ev.data as {
        type?: string
        payload?: { html?: string; text?: string; requestId?: string; forward?: boolean }
        html?: string
        text?: string
    } | null
    if (!d || typeof d.type !== 'string') return
    if (d.type === 'ONLYOFFICE_CURSOR_PROBE') {
        console.info('[DocumentHandler] 收到光标消息', d.payload)
        const p = d.payload as { event?: string; source?: string; doc?: Record<string, unknown> } | undefined
        if (p && typeof p.event === 'string') {
            emit('onlyofficeCursor', { event: p.event, source: p.source, doc: p.doc })
        }
        return
    }
    if (d.type === 'ONLYOFFICE_INSERT_HTML') {
        const html =
            typeof d.payload?.html === 'string' ? d.payload.html : typeof d.html === 'string' ? d.html : ''
        insertHtmlAtCursor(html)
        return
    }
    if (d.type === 'ONLYOFFICE_INSERT_TEXT') {
        const text =
            typeof d.payload?.text === 'string' ? d.payload.text : typeof d.text === 'string' ? d.text : ''
        insertTextAtCursor(text)
        return
    }
    if (d.type === 'ONLYOFFICE_GET_DOCUMENT_HTML') {
        const requestId = typeof d.payload?.requestId === 'string' ? d.payload.requestId : undefined
        const replyTo = (html: string | null) => {
            const reply = { type: 'ONLYOFFICE_GET_DOCUMENT_HTML_REPLY', payload: { requestId, html } }
            try {
                ;(ev.source as Window | null)?.postMessage?.(reply, '*')
            } catch (e) {
                console.warn('[OnlyOffice] GET_DOCUMENT_HTML reply', e)
            }
        }
        getDocumentHtml(replyTo)
        return
    }
    if (d.type === 'ONLYOFFICE_FIND_TEXT') {
        const requestId = typeof d.payload?.requestId === 'string' ? d.payload.requestId : undefined
        const text = typeof d.payload?.text === 'string' ? d.payload.text : ''
        const forward = d.payload?.forward !== false
        let found = false
        let error: string | undefined
        try {
            found = findTextInDocument(text, forward)
        } catch (e) {
            error = e instanceof Error ? e.message : String(e)
        }
        const reply = {
            type: 'ONLYOFFICE_FIND_TEXT_REPLY',
            payload: { requestId, found, error },
        }
        try {
            ;(ev.source as Window | null)?.postMessage?.(reply, '*')
        } catch {
            /* ignore */
        }
        try {
            if (window.parent && window.parent !== window) {
                window.parent.postMessage(reply, '*')
            }
        } catch {
            /* ignore */
        }
        return
    }
    if (d.type === 'ONLYOFFICE_REQUEST_SAVE_DOCUMENT') {
        const requestId =
            typeof d.payload?.requestId === 'string'
                ? d.payload.requestId
                : `save-${Date.now()}`
        void (async () => {
            try {
                const payload = await exportDocumentForSave(props.file.fileName)
                postSaveDocumentReply(ev.source, requestId, payload)
            } catch (e) {
                const msg = e instanceof Error ? e.message : String(e)
                console.warn('[OnlyOffice] exportDocumentForSave', msg)
                postSaveDocumentReply(ev.source, requestId, null, msg)
            }
        })()
    }
}

function getEditorAttachPair(api: Record<string, unknown>) {
    if (typeof api.attachEvent === 'function' && typeof api.detachEvent === 'function') {
        return {
            attach: (e: string, fn: () => void) => (api.attachEvent as (a: string, b: () => void) => void)(e, fn),
            detach: (e: string, fn: () => void) => (api.detachEvent as (a: string, b: () => void) => void)(e, fn),
        }
    }
    if (typeof api.q$ === 'function' && typeof api.yva === 'function') {
        return {
            attach: (e: string, fn: () => void) => (api.q$ as (a: string, b: () => void) => void)(e, fn),
            detach: (e: string, fn: () => void) => (api.yva as (a: string, b: () => void) => void)(e, fn),
        }
    }
    if (typeof api.Nca === 'function' && typeof api.DIa === 'function') {
        return {
            attach: (e: string, fn: () => void) => (api.Nca as (a: string, b: () => void) => void)(e, fn),
            detach: (e: string, fn: () => void) => (api.DIa as (a: string, b: () => void) => void)(e, fn),
        }
    }
    return null
}

function scheduleInnerCursorBridge() {
    clearInnerCursorBridge()
    const maxTries = 80
    let tries = 0
    innerBridgeTimer = setInterval(() => {
        tries++
        const iframe = document.querySelector('.editor-container iframe') as HTMLIFrameElement | null
        const w = iframe?.contentWindow as (Window & { Asc?: { editor?: Record<string, unknown> }; editor?: Record<string, unknown> }) | undefined
        const api = (w?.Asc?.editor ?? w?.editor) as Record<string, unknown> | undefined
        const pair = api ? getEditorAttachPair(api) : null
        if (!pair || !api) {
            if (tries >= maxTries) {
                clearInnerCursorBridge()
                console.warn('[cursor-bridge] 未在 iframe 内找到 Asc.editor（已重试', maxTries, '次）')
            }
            return
        }

        let throttle: ReturnType<typeof setTimeout> | null = null
        const fire = (ev: string) => {
            if (throttle) return
            throttle = setTimeout(() => {
                throttle = null
                postCursorProbe(ev, 'inner-bridge', api)
            }, 80)
        }
        const onMove = () => fire('asc_onCursorMove')
        const onSel = () => fire('asc_onSelectionEnd')
        const onModified = () => postDocumentModified(true)

        pair.attach('asc_onCursorMove', onMove)
        pair.attach('asc_onSelectionEnd', onSel)
        pair.attach('asc_onDocumentModifiedChanged', onModified)

        if (innerBridgeTimer) {
            clearInterval(innerBridgeTimer)
            innerBridgeTimer = null
        }

        innerBridgeCleanup = () => {
            try {
                pair.detach('asc_onCursorMove', onMove)
                pair.detach('asc_onSelectionEnd', onSel)
                pair.detach('asc_onDocumentModifiedChanged', onModified)
            } catch {
                /* ignore */
            }
        }

        console.info('[cursor-bridge] 已挂钩 Asc.editor（word/cell/slide 之一）')
        postCursorProbe('bridge_ready', 'inner-bridge', api)
    }, 250)
}

const editor = ref<any>(null)
const loading = ref(false)
let stopFileWatch: (() => void) | null = null

// 全局 media 映射对象
const media: { [key: string]: string } = {}

onMounted(async () => {
    window.addEventListener('message', onCursorPluginMessage)
    loading.value = true
    try {
        await initX2TScript()
        // 加载编辑器API
        await loadEditorApi()
        await initX2T()
        console.log('app has loading')
        loading.value = false

        stopFileWatch?.()
        stopFileWatch = watch(
            () => props.file.fileName,
            async () => {
                try {
                    await openFile()
                } catch (error) {
                    console.error('Error opening file:', error)
                    alert('文件打开失败，请检查文件格式')
                }
            },
            { immediate: true },
        )
    } catch (error) {
        console.error('Failed to initialize editor:', error)
        window.removeEventListener('message', onCursorPluginMessage)
        // 错误已在各函数中处理
    }
})
type EditorBinData = ArrayBuffer | Uint8Array | string

// 合并后的文件操作方法
async function handleDocumentOperation(options: {
    isNew: boolean
    fileName: string
    file?: File | null
}) {
    try {
        const { isNew, fileName, file } = options
        const fileType = fileName.split('.').pop() || ''
        getDocumentType(fileType)

        let binData: EditorBinData
        let media: Record<string, string> | undefined

        if (isNew) {
            const emptyBin = g_sEmpty_bin[`.${fileType}`]
            if (!emptyBin) {
                throw new Error(`不支持的文件类型: ${fileType}`)
            }
            binData = emptyBin
        } else {
            if (!file) throw new Error('无效的文件对象')
            const converted = await convertDocument(file)
            binData = converted.bin
            media = converted.media
        }

        createEditorInstance({
            fileName,
            fileType,
            binData,
            media,
        })
    } catch (error: any) {
        console.error('文档操作失败:', error)
        alert(`文档操作失败: ${error.message}`)
        throw error
    }
}

// 公共编辑器创建方法
function createEditorInstance(config: {
    fileName: string
    fileType: string
    binData: EditorBinData
    media?: Record<string, string>
}) {
    clearInnerCursorBridge()
    // 清理旧编辑器实例
    if (editor.value) {
        editor.value.destroyEditor()
        editor.value = null
    }

    const { fileName, fileType, binData, media } = config

    if (!window.DocsAPI) {
        throw new Error('OnlyOffice DocsAPI 未加载')
    }
    const readOnly = !!props.readonly
    editor.value = new window.DocsAPI.DocEditor('iframe', {
        document: {
            title: fileName,
            url: fileName, // 使用文件名作为标识
            fileType: fileType,
            permissions: {
                edit: !readOnly,
                download: true,
                print: true,
                // 只读也保持 review:true，避免 URL 出现 mode=view 导致 isEdit=false 触发 web-apps Header 的 userName 空指针
                review: true,
                comment: !readOnly,
                chat: false,
                protect: false,
            },
        },
        editorConfig: {
            lang: 'zh',
            user: {
                id: 'zb-office-viewer',
                name: '预览用户',
                firstname: '预览',
                lastname: '用户',
            },
            ...(readOnly
                ? {
                      coEditing: { mode: 'strict', change: false },
                  }
                : {
                      plugins: {
                          autostart: [CURSOR_PLUGIN_GUID],
                          pluginsData: [getCursorPluginConfigUrl()],
                      },
                  }),
            customization: {
                help: false,
                about: false,
                hideRightMenu: readOnly,
                comments: !readOnly,
                features: {
                    spellcheck: {
                        change: false,
                    },
                },
                anonymous: {
                    request: false,
                    label: '预览用户',
                },
            },
        },
        events: {
            onAppReady: () => {
                // 设置媒体资源
                if (media) {
                    editor.value.sendCommand({
                        command: 'asc_setImageUrls',
                        data: { urls: media },
                    })
                }

                // 加载文档内容
                editor.value.sendCommand({
                    command: 'asc_openDocument',
                    data: { buf: binData },
                })
            },
            onDocumentReady: () => {
                console.log('文档加载完成:', fileName)
                if (readOnly) {
                    const api = getEditorApiFromDom()
                    if (api && typeof api.asc_setViewMode === 'function') {
                        try {
                            ;(api.asc_setViewMode as (this: unknown, v: boolean) => void).call(api, true)
                        } catch (e) {
                            console.warn('[OnlyOffice] asc_setViewMode', e)
                        }
                    }
                } else {
                    scheduleInnerCursorBridge()
                }
                const readyMsg = { type: 'ONLYOFFICE_DOCUMENT_READY' }
                window.postMessage(readyMsg, '*')
                try {
                    if (window.parent && window.parent !== window) {
                        window.parent.postMessage(readyMsg, '*')
                    }
                } catch {
                    /* ignore */
                }
            },
            onSave: readOnly ? undefined : handleSaveDocument,
            // writeFile
            // todo writeFile 当外部粘贴图片时候处理
            writeFile: handleWriteFile,
        },
    })
}

// 修改后的openFile方法
async function openFile() {
    const { fileName, file } = props.file

    await handleDocumentOperation({
        isNew: !file, // 根据是否存在file判断是否新建
        fileName,
        file,
    })
}

function loadEditorApi(): Promise<void> {
    return new Promise((resolve, reject) => {
        // 检查是否已加载
        if (window.DocsAPI) {
            resolve()
            return
        }

        // 加载编辑器API
        const script = document.createElement('script')
        script.src = './web-apps/apps/api/documents/api.js'
        script.onload = () => resolve()
        script.onerror = (error) => {
            console.error('Failed to load OnlyOffice API:', error)
            alert('无法加载编辑器组件。请确保已正确安装 OnlyOffice API。')
            reject(error)
        }
        document.head.appendChild(script)
    })
}

interface SaveEvent {
    data?: {
        data?: { data?: string }
        option?: { outputformat?: string }
    }
}

async function handleSaveDocument(event: SaveEvent) {
    const savePayload = event.data?.data
    const option = event.data?.option
    let saveError: string | undefined

    if (savePayload?.data && option?.outputformat) {
        try {
            const fmt = option.outputformat as unknown as keyof typeof c_oAscFileType2
            const converted = await convertBinToDocument(
                Uint8Array.from(atob(savePayload.data), (c) => c.charCodeAt(0)),
                props.file.fileName,
                c_oAscFileType2[fmt],
            )
            flushPendingSaveReply(undefined, {
                fileName: converted.fileName,
                bytes: converted.data,
            })
        } catch (e) {
            saveError = e instanceof Error ? e.message : String(e)
            flushPendingSaveReply(undefined, null, saveError)
        }
    } else if (pendingSaveReplies.size > 0) {
        flushPendingSaveReply(undefined, null, '保存数据为空')
    }

    editor.value.sendCommand({
        command: 'asc_onSaveCallback',
        data: { err_code: saveError ? 1 : 0 },
    })
}

// 辅助函数：将base64转为Blob
function dataURItoBlob(dataURI: string): Blob {
    // 从base64字符串中提取数据部分
    const byteString = atob(dataURI.split(',')[1])

    // 创建ArrayBuffer
    const ab = new ArrayBuffer(byteString.length)
    const ia = new Uint8Array(ab)

    for (let i = 0; i < byteString.length; i++) {
        ia[i] = byteString.charCodeAt(i)
    }

    return new Blob([ab])
}

/**
 * 处理文件写入请求（主要用于处理粘贴的图片）
 * @param event - OnlyOffice 编辑器的文件写入事件
 */
function handleWriteFile(event: any) {
    // debugger
    try {
        console.log('Write file event:', event)

        const { data: eventData } = event
        if (!eventData) {
            console.warn('No data provided in writeFile event')
            return
        }

        const {
            data: imageData, // Uint8Array 图片数据
            file: fileName, // 文件名，如 "display8image-174799443357-0.png"
            target, // 目标对象，包含 frameOrigin 等信息
        } = eventData

        // 验证数据
        if (!imageData || !(imageData instanceof Uint8Array)) {
            throw new Error('Invalid image data: expected Uint8Array')
        }

        if (!fileName || typeof fileName !== 'string') {
            throw new Error('Invalid file name')
        }

        // 从文件名中提取扩展名
        const fileExtension = fileName.split('.').pop()?.toLowerCase() || 'png'
        const mimeType = getMimeTypeFromExtension(fileExtension)

        // 创建 Blob 对象
        const blob = new Blob([imageData as BlobPart], { type: mimeType })

        // 创建对象 URL
        const objectUrl = URL.createObjectURL(blob)
        // 将图片设置为base64url
        //  const base64Url = `data:${mimeType};base64,${btoa(String.fromCharCode(...imageData))}`;
        // 将图片URL添加到媒体映射中，使用原始文件名作为key
        media[`media/${fileName}`] = objectUrl
        editor.value.sendCommand({
            command: 'asc_setImageUrls',
            data: {
                urls: media,
            },
        })

        editor.value.sendCommand({
            command: 'asc_writeFileCallback',
            data: {
                // 图片base64
                path: objectUrl,
                imgName: fileName,
            },
        })
        console.log(`Successfully processed image: ${fileName}, URL: ${media}`)
    } catch (error: unknown) {
        console.error('Error handling writeFile:', error)
        const errMsg = error instanceof Error ? error.message : String(error)

        if (editor.value && typeof editor.value.sendCommand === 'function') {
            editor.value.sendCommand({
                command: 'asc_writeFileCallback',
                data: {
                    success: false,
                    error: errMsg,
                },
            })
        }

        if (event.callback && typeof event.callback === 'function') {
            event.callback({
                success: false,
                error: errMsg,
            })
        }
    }
}

/**
 * 根据文件扩展名获取 MIME 类型
 * @param extension - 文件扩展名
 * @returns string - MIME 类型
 */
function getMimeTypeFromExtension(extension: string): string {
    const mimeMap: { [key: string]: string } = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        gif: 'image/gif',
        bmp: 'image/bmp',
        webp: 'image/webp',
        svg: 'image/svg+xml',
        ico: 'image/x-icon',
        tiff: 'image/tiff',
        tif: 'image/tiff',
    }

    return mimeMap[extension?.toLowerCase()] || 'image/png'
}

// 组件卸载时清理对象 URL（须在 setup 顶层注册，不可放在 onMounted 的 await 之后）
onBeforeUnmount(() => {
    stopFileWatch?.()
    stopFileWatch = null
    clearInnerCursorBridge()
    window.removeEventListener('message', onCursorPluginMessage)
    // 清理媒体资源的对象 URL
    Object.values(media).forEach((url) => {
        if (typeof url === 'string' && url.startsWith('blob:')) {
            URL.revokeObjectURL(url)
        }
    })

    // 清理编辑器资源
    if (editor.value) {
        if (typeof editor.value.destroyEditor === 'function') {
            editor.value.destroyEditor()
        }
    }
})
</script>

<style scoped>
.editor-container {
    width: 100%;
    height: 100vh;
}

#iframe {
    width: 100%;
    height: 100%;
}
</style>

