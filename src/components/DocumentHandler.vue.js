import { onMounted, onBeforeUnmount, ref, watch } from 'vue';
import { getDocumentType } from '@/utils/util';
import { g_sEmpty_bin } from '@/utils/empty_bin';
// @ts-ignore
import { initX2TScript, initX2T, convertDocument, convertBinToDocument, c_oAscFileType2, oAscFileType, } from '@/utils/x2t';
const X2T = ref(null);
const props = defineProps();
const emit = defineEmits();
const CURSOR_PLUGIN_GUID = 'asc.{6C757CD6-7A91-4B0F-9C2D-1E3F4A5B6C7D}';
function getCursorPluginConfigUrl() {
    try {
        let base = import.meta.env.BASE_URL || '/';
        if (!base.endsWith('/'))
            base += '/';
        return new URL('onlyoffice-cursor-plugin/config.json', new URL(base, window.location.origin)).href;
    }
    catch {
        return `${window.location.origin}/onlyoffice-cursor-plugin/config.json`;
    }
}
/** 本地 web-apps 常未真正加载自定义插件，改为在编辑器 iframe 内订阅 Asc.editor 事件（同源） */
let innerBridgeTimer = null;
let innerBridgeCleanup = null;
function clearInnerCursorBridge() {
    if (innerBridgeTimer) {
        clearInterval(innerBridgeTimer);
        innerBridgeTimer = null;
    }
    if (innerBridgeCleanup) {
        innerBridgeCleanup();
        innerBridgeCleanup = null;
    }
}
/** 与任意父页面约定：type === ONLYOFFICE_CURSOR_PROBE，payload 含 event / source / doc */
function postCursorProbe(event, source, api) {
    const doc = api ? collectEditorCursorSnapshot(api) : undefined;
    const msg = { type: 'ONLYOFFICE_CURSOR_PROBE', payload: { event, source, doc } };
    window.postMessage(msg, '*');
    try {
        if (window.top && window.top !== window) {
            window.top.postMessage(msg, '*');
        }
    }
    catch {
        /* 跨域 top 不可写时忽略 */
    }
}
/** 使用编辑器内置 pluginMethod（与插件同源），读取选区文本 / 类型；并读取插入点/选区在版式中的坐标（非浏览器像素） */
function serializeAnchorRect(pos) {
    if (!pos || typeof pos !== 'object')
        return null;
    const p = pos;
    const ax = p.asc_getX;
    const ay = p.asc_getY;
    const aw = p.asc_getWidth;
    const ah = p.asc_getHeight;
    if (typeof ax === 'function' &&
        typeof ay === 'function' &&
        typeof aw === 'function' &&
        typeof ah === 'function') {
        return {
            x: ax.call(pos),
            y: ay.call(pos),
            width: aw.call(pos),
            height: ah.call(pos),
        };
    }
    return null;
}
function collectEditorCursorSnapshot(api) {
    const out = {};
    try {
        const getSel = api.pluginMethod_GetSelectedText;
        const getType = api.pluginMethod_GetSelectionType;
        if (typeof getSel === 'function') {
            const selectedText = getSel.call(api, {
                Numbering: true,
                Math: true,
            });
            out.selectedText = selectedText;
            out.selectedLength =
                typeof selectedText === 'string'
                    ? selectedText.length
                    : Array.isArray(selectedText)
                        ? JSON.stringify(selectedText).length
                        : 0;
        }
        if (typeof getType === 'function') {
            out.selectionType = getType.call(api);
        }
        if (api.cn !== undefined)
            out.editorCn = api.cn;
        const len = typeof out.selectedLength === 'number' ? out.selectedLength : 0;
        out.isCollapsed = len === 0;
        const getAnchor = api.asc_getAnchorPosition;
        if (typeof getAnchor === 'function') {
            try {
                const pos = getAnchor.call(api);
                const rect = serializeAnchorRect(pos);
                if (rect)
                    out.anchor = rect;
            }
            catch (e) {
                out.anchorError = String(e);
            }
        }
        const getBounds = api.asc_GetSelectionBounds;
        if (typeof getBounds === 'function') {
            try {
                out.selectionBounds = getBounds.call(api);
            }
            catch (e) {
                out.selectionBoundsError = String(e);
            }
        }
    }
    catch (e) {
        out.snapshotError = String(e);
    }
    return out;
}
function getEditorApiFromDom() {
    const iframe = document.querySelector('.editor-container iframe');
    const w = iframe?.contentWindow;
    const api = (w?.Asc?.editor ?? w?.editor);
    return api ?? null;
}
/** 在当前光标处插入纯文本（Word/表格等内置支持 pluginMethod_InputText 时可用） */
function insertTextAtCursor(text) {
    const api = getEditorApiFromDom();
    if (!api || typeof api.pluginMethod_InputText !== 'function') {
        console.warn('[OnlyOffice] pluginMethod_InputText 不可用（可能未打开文档或非文本模式）');
        return false;
    }
    try {
        ;
        api.pluginMethod_InputText.call(api, String(text ?? ''), '');
        return true;
    }
    catch (e) {
        console.error('[OnlyOffice] insertTextAtCursor 失败', e);
        return false;
    }
}
/** 在当前光标处粘贴 HTML 片段（需 Word 可编辑且 pluginMethod_PasteHtml 可用） */
function insertHtmlAtCursor(html) {
    const api = getEditorApiFromDom();
    if (!api || typeof api.pluginMethod_PasteHtml !== 'function') {
        console.warn('[OnlyOffice] pluginMethod_PasteHtml 不可用');
        return false;
    }
    try {
        ;
        api.pluginMethod_PasteHtml.call(api, String(html ?? ''));
        return true;
    }
    catch (e) {
        console.error('[OnlyOffice] insertHtmlAtCursor 失败', e);
        return false;
    }
}
/** 整篇 HTML：编辑器 iframe 内 asc_nativeGetHtml（当前环境无 ToHtml 时用此路径） */
function getDocumentHtmlSync() {
    const api = getEditorApiFromDom();
    if (!api)
        return null;
    if (typeof api.asc_nativeGetHtml === 'function') {
        try {
            const html = api.asc_nativeGetHtml.call(api);
            if (typeof html === 'string' && html.trim()) {
                console.info('[OnlyOffice] getDocumentHtml: asc_nativeGetHtml', html.length);
                return html;
            }
        }
        catch (e) {
            console.warn('[OnlyOffice] asc_nativeGetHtml', e);
        }
    }
    if (typeof api.pluginMethod_GetFileHTML === 'function') {
        try {
            const html = api.pluginMethod_GetFileHTML.call(api);
            if (typeof html === 'string' && html.trim()) {
                console.info('[OnlyOffice] getDocumentHtml: pluginMethod_GetFileHTML', html.length);
                return html;
            }
        }
        catch (e) {
            console.warn('[OnlyOffice] pluginMethod_GetFileHTML', e);
        }
    }
    console.warn('[OnlyOffice] getDocumentHtml: asc_nativeGetHtml / GetFileHTML 均不可用');
    return null;
}
function getDocumentHtml(callback) {
    callback(getDocumentHtmlSync());
}
const __VLS_exposed = { insertTextAtCursor, insertHtmlAtCursor, getDocumentHtml, getEditorApiFromDom };
defineExpose(__VLS_exposed);
/** 父页 REQUEST_SAVE 等待 onSave 完成后回传 docx */
const pendingSaveReplies = new Map();
function postSaveDocumentReply(target, requestId, payload, error) {
    const reply = {
        type: 'ONLYOFFICE_SAVE_DOCUMENT_REPLY',
        payload: { requestId, fileName: payload?.fileName, base64: payload?.base64, error },
    };
    try {
        ;
        target?.postMessage?.(reply, '*');
    }
    catch {
        /* ignore */
    }
    try {
        if (window.parent && window.parent !== window) {
            window.parent.postMessage(reply, '*');
        }
    }
    catch {
        /* ignore */
    }
}
function flushPendingSaveReply(requestId, payload, error) {
    if (requestId) {
        const fn = pendingSaveReplies.get(requestId);
        if (fn) {
            pendingSaveReplies.delete(requestId);
            fn(payload, error);
            return;
        }
    }
    for (const [id, fn] of pendingSaveReplies.entries()) {
        pendingSaveReplies.delete(id);
        fn(payload, error);
        break;
    }
}
function uint8ToBase64(bytes) {
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
}
async function blobToSavePayload(blob, fileName) {
    const buf = await blob.arrayBuffer();
    return { fileName, base64: uint8ToBase64(new Uint8Array(buf)) };
}
async function resolveDownloadDataToPayload(data, fileName) {
    if (data.startsWith('http://') || data.startsWith('https://') || data.startsWith('blob:')) {
        const res = await fetch(data);
        if (!res.ok)
            throw new Error(`拉取导出文件失败: ${res.status}`);
        return blobToSavePayload(await res.blob(), fileName);
    }
    return { fileName, base64: data };
}
function hijackApiDownloadCallback(api, fileName, timeoutMs) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('导出超时')), timeoutMs);
        const prevAha = api.Aha;
        api.Aha = (result) => {
            clearTimeout(timer);
            if (typeof prevAha === 'function')
                api.Aha = prevAha;
            const r = result;
            if (r?.status === 'ok' && r?.data && r.data !== 'error') {
                void resolveDownloadDataToPayload(r.data, fileName).then(resolve).catch(reject);
            }
            else {
                reject(new Error('编辑器导出失败'));
            }
        };
    });
}
function dispatchEditorSaveShortcut() {
    const iframe = document.querySelector('.editor-container iframe');
    const doc = iframe?.contentWindow?.document;
    if (!doc)
        return false;
    const opts = { key: 's', code: 'KeyS', ctrlKey: true, bubbles: true, cancelable: true };
    doc.dispatchEvent(new KeyboardEvent('keydown', opts));
    doc.dispatchEvent(new KeyboardEvent('keyup', opts));
    return true;
}
function triggerEditorSaveCommands() {
    const api = getEditorApiFromDom();
    if (api && typeof api.asc_Save === 'function') {
        try {
            ;
            api.asc_Save.call(api);
        }
        catch (e) {
            console.warn('[OnlyOffice] asc_Save failed', e);
        }
    }
    if (editor.value && typeof editor.value.sendCommand === 'function') {
        editor.value.sendCommand({ command: 'asc_Save', data: {} });
    }
    dispatchEditorSaveShortcut();
}
/** 通过 DocEditor sendCommand 触发 onSave → convertBinToDocument（本地嵌入最稳） */
function exportViaEditorOnSave(fileName, timeoutMs) {
    return new Promise((resolve, reject) => {
        if (!editor.value || typeof editor.value.sendCommand !== 'function') {
            reject(new Error('编辑器未就绪'));
            return;
        }
        const requestId = `save-${Date.now()}`;
        const timer = setTimeout(() => {
            pendingSaveReplies.delete(requestId);
            reject(new Error('onSave 未触发'));
        }, timeoutMs);
        pendingSaveReplies.set(requestId, (payload, err) => {
            clearTimeout(timer);
            if (err || !payload)
                reject(new Error(err || '保存数据为空'));
            else
                resolve(payload);
        });
        try {
            editor.value.sendCommand({
                command: 'asc_DownloadAs',
                data: { format: oAscFileType.DOCX, fileType: oAscFileType.DOCX },
            });
        }
        catch (e) {
            clearTimeout(timer);
            pendingSaveReplies.delete(requestId);
            reject(e instanceof Error ? e : new Error(String(e)));
        }
    });
}
/** 导出当前文档为 docx（供父页自动保存；本地壳页不依赖用户手动点保存） */
async function exportDocumentForSave(fileName) {
    try {
        return await exportViaEditorOnSave(fileName, 25_000);
    }
    catch (e) {
        console.warn('[OnlyOffice] exportViaEditorOnSave', e);
    }
    const api = getEditorApiFromDom();
    const ext = (fileName.split('.').pop() || 'docx').toUpperCase();
    if (api && typeof api.pluginMethod_GetFileToDownload === 'function') {
        const pending = hijackApiDownloadCallback(api, fileName, 30_000);
        try {
            ;
            api.pluginMethod_GetFileToDownload.call(api, ext);
            return await pending;
        }
        catch (e) {
            console.warn('[OnlyOffice] pluginMethod_GetFileToDownload failed', e);
        }
    }
    if (api && typeof api.asc_DownloadAs === 'function') {
        const pending = hijackApiDownloadCallback(api, fileName, 30_000);
        try {
            ;
            api.asc_DownloadAs.call(api);
            return await pending;
        }
        catch (e) {
            console.warn('[OnlyOffice] asc_DownloadAs failed', e);
        }
    }
    return new Promise((resolve, reject) => {
        const requestId = `save-${Date.now()}`;
        const timer = setTimeout(() => {
            pendingSaveReplies.delete(requestId);
            reject(new Error('onSave 未触发，请确认 onlyoffice 壳页已重新编译'));
        }, 20_000);
        pendingSaveReplies.set(requestId, (payload, err) => {
            clearTimeout(timer);
            if (err || !payload)
                reject(new Error(err || '保存数据为空'));
            else
                resolve(payload);
        });
        triggerEditorSaveCommands();
    });
}
/** 同源父页用 ref；最外层 iframe 向本窗口 postMessage：INSERT_* / GET_DOCUMENT_HTML（回复 ONLYOFFICE_GET_DOCUMENT_HTML_REPLY） */
function onCursorPluginMessage(ev) {
    const d = ev.data;
    if (!d || typeof d.type !== 'string')
        return;
    if (d.type === 'ONLYOFFICE_CURSOR_PROBE') {
        console.info('[DocumentHandler] 收到光标消息', d.payload);
        const p = d.payload;
        if (p && typeof p.event === 'string') {
            emit('onlyofficeCursor', { event: p.event, source: p.source, doc: p.doc });
        }
        return;
    }
    if (d.type === 'ONLYOFFICE_INSERT_HTML') {
        const html = typeof d.payload?.html === 'string' ? d.payload.html : typeof d.html === 'string' ? d.html : '';
        insertHtmlAtCursor(html);
        return;
    }
    if (d.type === 'ONLYOFFICE_INSERT_TEXT') {
        const text = typeof d.payload?.text === 'string' ? d.payload.text : typeof d.text === 'string' ? d.text : '';
        insertTextAtCursor(text);
        return;
    }
    if (d.type === 'ONLYOFFICE_GET_DOCUMENT_HTML') {
        const requestId = typeof d.payload?.requestId === 'string' ? d.payload.requestId : undefined;
        const replyTo = (html) => {
            const reply = { type: 'ONLYOFFICE_GET_DOCUMENT_HTML_REPLY', payload: { requestId, html } };
            try {
                ;
                ev.source?.postMessage?.(reply, '*');
            }
            catch (e) {
                console.warn('[OnlyOffice] GET_DOCUMENT_HTML reply', e);
            }
        };
        getDocumentHtml(replyTo);
        return;
    }
    if (d.type === 'ONLYOFFICE_REQUEST_SAVE_DOCUMENT') {
        const requestId = typeof d.payload?.requestId === 'string'
            ? d.payload.requestId
            : `save-${Date.now()}`;
        void (async () => {
            try {
                const payload = await exportDocumentForSave(props.file.fileName);
                postSaveDocumentReply(ev.source, requestId, payload);
            }
            catch (e) {
                const msg = e instanceof Error ? e.message : String(e);
                console.error('[OnlyOffice] export for save failed', e);
                postSaveDocumentReply(ev.source, requestId, null, msg);
            }
        })();
    }
}
function getEditorAttachPair(api) {
    if (typeof api.attachEvent === 'function' && typeof api.detachEvent === 'function') {
        return {
            attach: (e, fn) => api.attachEvent(e, fn),
            detach: (e, fn) => api.detachEvent(e, fn),
        };
    }
    if (typeof api.q$ === 'function' && typeof api.yva === 'function') {
        return {
            attach: (e, fn) => api.q$(e, fn),
            detach: (e, fn) => api.yva(e, fn),
        };
    }
    if (typeof api.Nca === 'function' && typeof api.DIa === 'function') {
        return {
            attach: (e, fn) => api.Nca(e, fn),
            detach: (e, fn) => api.DIa(e, fn),
        };
    }
    return null;
}
function scheduleInnerCursorBridge() {
    clearInnerCursorBridge();
    const maxTries = 80;
    let tries = 0;
    innerBridgeTimer = setInterval(() => {
        tries++;
        const iframe = document.querySelector('.editor-container iframe');
        const w = iframe?.contentWindow;
        const api = (w?.Asc?.editor ?? w?.editor);
        const pair = api ? getEditorAttachPair(api) : null;
        if (!pair || !api) {
            if (tries >= maxTries) {
                clearInnerCursorBridge();
                console.warn('[cursor-bridge] 未在 iframe 内找到 Asc.editor（已重试', maxTries, '次）');
            }
            return;
        }
        let throttle = null;
        const fire = (ev) => {
            if (throttle)
                return;
            throttle = setTimeout(() => {
                throttle = null;
                postCursorProbe(ev, 'inner-bridge', api);
            }, 80);
        };
        const onMove = () => fire('asc_onCursorMove');
        const onSel = () => fire('asc_onSelectionEnd');
        pair.attach('asc_onCursorMove', onMove);
        pair.attach('asc_onSelectionEnd', onSel);
        if (innerBridgeTimer) {
            clearInterval(innerBridgeTimer);
            innerBridgeTimer = null;
        }
        innerBridgeCleanup = () => {
            try {
                pair.detach('asc_onCursorMove', onMove);
                pair.detach('asc_onSelectionEnd', onSel);
            }
            catch {
                /* ignore */
            }
        };
        console.info('[cursor-bridge] 已挂钩 Asc.editor（word/cell/slide 之一）');
        postCursorProbe('bridge_ready', 'inner-bridge', api);
    }, 250);
}
const editor = ref(null);
const loading = ref(false);
// 全局 media 映射对象
const media = {};
onMounted(async () => {
    window.addEventListener('message', onCursorPluginMessage);
    loading.value = true;
    try {
        await initX2TScript();
        // 加载编辑器API
        await loadEditorApi();
        await initX2T();
        console.log('app has loading');
        loading.value = false;
        // 页面初始化后，使用 watchEffect 监听 props.file 并执行 openFile
        // 添加props.file监听
        const stopWatch = watch(() => props.file.fileName, async () => {
            try {
                await openFile();
            }
            catch (error) {
                console.error('Error opening file:', error);
                alert('文件打开失败，请检查文件格式');
            }
        }, { immediate: true });
        // 组件卸载时停止监听
        onBeforeUnmount(stopWatch);
    }
    catch (error) {
        console.error('Failed to initialize editor:', error);
        window.removeEventListener('message', onCursorPluginMessage);
        // 错误已在各函数中处理
    }
});
// 合并后的文件操作方法
async function handleDocumentOperation(options) {
    try {
        const { isNew, fileName, file } = options;
        const fileType = fileName.split('.').pop() || '';
        getDocumentType(fileType);
        let binData;
        let media;
        if (isNew) {
            const emptyBin = g_sEmpty_bin[`.${fileType}`];
            if (!emptyBin) {
                throw new Error(`不支持的文件类型: ${fileType}`);
            }
            binData = emptyBin;
        }
        else {
            if (!file)
                throw new Error('无效的文件对象');
            const converted = await convertDocument(file);
            binData = converted.bin;
            media = converted.media;
        }
        createEditorInstance({
            fileName,
            fileType,
            binData,
            media,
        });
    }
    catch (error) {
        console.error('文档操作失败:', error);
        alert(`文档操作失败: ${error.message}`);
        throw error;
    }
}
// 公共编辑器创建方法
function createEditorInstance(config) {
    clearInnerCursorBridge();
    // 清理旧编辑器实例
    if (editor.value) {
        editor.value.destroyEditor();
        editor.value = null;
    }
    const { fileName, fileType, binData, media } = config;
    if (!window.DocsAPI) {
        throw new Error('OnlyOffice DocsAPI 未加载');
    }
    editor.value = new window.DocsAPI.DocEditor('iframe', {
        document: {
            title: fileName,
            url: fileName, // 使用文件名作为标识
            fileType: fileType,
            permissions: {
                edit: true,
                chat: false,
                protect: false,
            },
        },
        editorConfig: {
            lang: 'zh',
            plugins: {
                autostart: [CURSOR_PLUGIN_GUID],
                pluginsData: [getCursorPluginConfigUrl()],
            },
            customization: {
                help: false,
                about: false,
                hideRightMenu: true,
                features: {
                    spellcheck: {
                        change: false,
                    },
                },
                anonymous: {
                    request: false,
                    label: 'Guest',
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
                    });
                }
                // 加载文档内容
                editor.value.sendCommand({
                    command: 'asc_openDocument',
                    data: { buf: binData },
                });
            },
            onDocumentReady: () => {
                console.log('文档加载完成:', fileName);
                scheduleInnerCursorBridge();
                const readyMsg = { type: 'ONLYOFFICE_DOCUMENT_READY' };
                window.postMessage(readyMsg, '*');
                try {
                    if (window.parent && window.parent !== window) {
                        window.parent.postMessage(readyMsg, '*');
                    }
                }
                catch {
                    /* ignore */
                }
            },
            onSave: handleSaveDocument,
            // writeFile
            // todo writeFile 当外部粘贴图片时候处理
            writeFile: handleWriteFile,
        },
    });
}
// 修改后的openFile方法
async function openFile() {
    const { fileName, file } = props.file;
    await handleDocumentOperation({
        isNew: !file, // 根据是否存在file判断是否新建
        fileName,
        file,
    });
}
function loadEditorApi() {
    return new Promise((resolve, reject) => {
        // 检查是否已加载
        if (window.DocsAPI) {
            resolve();
            return;
        }
        // 加载编辑器API
        const script = document.createElement('script');
        script.src = './web-apps/apps/api/documents/api.js';
        script.onload = () => resolve();
        script.onerror = (error) => {
            console.error('Failed to load OnlyOffice API:', error);
            alert('无法加载编辑器组件。请确保已正确安装 OnlyOffice API。');
            reject(error);
        };
        document.head.appendChild(script);
    });
}
async function handleSaveDocument(event) {
    const savePayload = event.data?.data;
    const option = event.data?.option;
    let saveError;
    if (savePayload?.data && option?.outputformat) {
        try {
            const fmt = option.outputformat;
            const converted = await convertBinToDocument(Uint8Array.from(atob(savePayload.data), (c) => c.charCodeAt(0)), props.file.fileName, c_oAscFileType2[fmt]);
            flushPendingSaveReply(undefined, {
                fileName: converted.fileName,
                base64: uint8ToBase64(converted.data),
            });
        }
        catch (e) {
            saveError = e instanceof Error ? e.message : String(e);
            flushPendingSaveReply(undefined, null, saveError);
        }
    }
    else if (pendingSaveReplies.size > 0) {
        flushPendingSaveReply(undefined, null, '保存数据为空');
    }
    editor.value.sendCommand({
        command: 'asc_onSaveCallback',
        data: { err_code: saveError ? 1 : 0 },
    });
}
// 辅助函数：将base64转为Blob
function dataURItoBlob(dataURI) {
    // 从base64字符串中提取数据部分
    const byteString = atob(dataURI.split(',')[1]);
    // 创建ArrayBuffer
    const ab = new ArrayBuffer(byteString.length);
    const ia = new Uint8Array(ab);
    for (let i = 0; i < byteString.length; i++) {
        ia[i] = byteString.charCodeAt(i);
    }
    return new Blob([ab]);
}
/**
 * 处理文件写入请求（主要用于处理粘贴的图片）
 * @param event - OnlyOffice 编辑器的文件写入事件
 */
function handleWriteFile(event) {
    // debugger
    try {
        console.log('Write file event:', event);
        const { data: eventData } = event;
        if (!eventData) {
            console.warn('No data provided in writeFile event');
            return;
        }
        const { data: imageData, // Uint8Array 图片数据
        file: fileName, // 文件名，如 "display8image-174799443357-0.png"
        target, // 目标对象，包含 frameOrigin 等信息
         } = eventData;
        // 验证数据
        if (!imageData || !(imageData instanceof Uint8Array)) {
            throw new Error('Invalid image data: expected Uint8Array');
        }
        if (!fileName || typeof fileName !== 'string') {
            throw new Error('Invalid file name');
        }
        // 从文件名中提取扩展名
        const fileExtension = fileName.split('.').pop()?.toLowerCase() || 'png';
        const mimeType = getMimeTypeFromExtension(fileExtension);
        // 创建 Blob 对象
        const blob = new Blob([imageData], { type: mimeType });
        // 创建对象 URL
        const objectUrl = URL.createObjectURL(blob);
        // 将图片设置为base64url
        //  const base64Url = `data:${mimeType};base64,${btoa(String.fromCharCode(...imageData))}`;
        // 将图片URL添加到媒体映射中，使用原始文件名作为key
        media[`media/${fileName}`] = objectUrl;
        editor.value.sendCommand({
            command: 'asc_setImageUrls',
            data: {
                urls: media,
            },
        });
        editor.value.sendCommand({
            command: 'asc_writeFileCallback',
            data: {
                // 图片base64
                path: objectUrl,
                imgName: fileName,
            },
        });
        console.log(`Successfully processed image: ${fileName}, URL: ${media}`);
    }
    catch (error) {
        console.error('Error handling writeFile:', error);
        const errMsg = error instanceof Error ? error.message : String(error);
        if (editor.value && typeof editor.value.sendCommand === 'function') {
            editor.value.sendCommand({
                command: 'asc_writeFileCallback',
                data: {
                    success: false,
                    error: errMsg,
                },
            });
        }
        if (event.callback && typeof event.callback === 'function') {
            event.callback({
                success: false,
                error: errMsg,
            });
        }
    }
}
/**
 * 根据文件扩展名获取 MIME 类型
 * @param extension - 文件扩展名
 * @returns string - MIME 类型
 */
function getMimeTypeFromExtension(extension) {
    const mimeMap = {
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
    };
    return mimeMap[extension?.toLowerCase()] || 'image/png';
}
// 组件卸载时清理对象 URL
onBeforeUnmount(() => {
    clearInnerCursorBridge();
    window.removeEventListener('message', onCursorPluginMessage);
    // 清理媒体资源的对象 URL
    Object.values(media).forEach((url) => {
        if (typeof url === 'string' && url.startsWith('blob:')) {
            URL.revokeObjectURL(url);
        }
    });
    // 清理编辑器资源
    if (editor.value) {
        if (typeof editor.value.destroyEditor === 'function') {
            editor.value.destroyEditor();
        }
    }
});
debugger; /* PartiallyEnd: #3632/scriptSetup.vue */
const __VLS_ctx = {};
let __VLS_components;
let __VLS_directives;
// CSS variable injection 
// CSS variable injection end 
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "editor-container" },
    'element-loading-text': "Loading...",
});
__VLS_asFunctionalDirective(__VLS_directives.vLoaing)(null, { ...__VLS_directiveBindingRestFields, value: (__VLS_ctx.loading) }, null, null);
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    id: "iframe",
});
/** @type {__VLS_StyleScopedClasses['editor-container']} */ ;
var __VLS_dollars;
const __VLS_self = (await import('vue')).defineComponent({
    setup() {
        return {
            loading: loading,
        };
    },
    __typeEmits: {},
    __typeProps: {},
});
export default (await import('vue')).defineComponent({
    setup() {
        return {
            ...__VLS_exposed,
        };
    },
    __typeEmits: {},
    __typeProps: {},
});
; /* PartiallyEnd: #4569/main.vue */
//# sourceMappingURL=DocumentHandler.vue.js.map