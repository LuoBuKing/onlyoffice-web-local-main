import { FolderOpened } from '@element-plus/icons-vue';
import { onMounted, ref } from 'vue';
import DocumentHandler from '../components/DocumentHandler.vue';
import { useRoute } from 'vue-router';
import { ElLoading } from 'element-plus';
const showCreateDialog = ref(false);
const selectedFile = ref(null);
const documentHandler = ref(null);
const docmentObj = ref(null);
/** doc：selectedText、isCollapsed、anchor(版式 x,y,w,h)、selectionBounds；插入用 documentHandler.insertTextAtCursor */
function onOnlyofficeCursor(payload) {
    console.log('[OnlyOffice 光标探针]', payload.event, payload.source ?? '', payload.doc ?? '(no doc)');
}
const onCreateNew = (ext) => {
    docmentObj.value = {
        fileName: '新建文档' + ext,
        file: null,
    };
    showCreateDialog.value = false;
};
const onOpenDocument = async () => {
    // 创建文件选择器，选择Office文档
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx,.xlsx,.pptx,.doc,.xls,.ppt';
    input.onchange = (event) => {
        const file = event.target.files?.[0];
        if (file) {
            showCreateDialog.value = false;
            docmentObj.value = {
                fileName: file.name,
                file: file,
            };
        }
    };
    input.click();
};
// 页面初始化后根据路由地址获取文件 并自动打开
async function initFileUrl() {
    const route = useRoute();
    const url = route.query.url;
    const filenameParam = route.query.filename;
    if (!url) {
        console.warn('未提供文件 URL');
        return;
    }
    const laodingInstance = ElLoading.service({
        lock: true,
        text: 'Loading',
        background: 'rgba(0, 0, 0, 0.7)',
    });
    try {
        const res = await fetch(url);
        if (!res.ok)
            throw new Error('文件请求失败');
        laodingInstance.close();
        const blob = await res.blob();
        let fileName = '';
        // 1. 从 query 参数获取 filename
        if (filenameParam) {
            fileName = filenameParam;
        }
        // 2. 如果没有 filename 参数，尝试从 URL 末尾解析
        if (!fileName) {
            const match = decodeURIComponent(url).match(/\/([^\/?#]+)$/);
            if (match && match[1].includes('.')) {
                fileName = match[1];
            }
        }
        // 3. 如果 URL 也解析失败，尝试从 Content-Disposition 响应头获取
        if (!fileName) {
            const disposition = res.headers.get('Content-Disposition');
            if (disposition) {
                const match = disposition.match(/filename\*=UTF-8''(.+)|filename="?([^"]+)"?/);
                if (match) {
                    fileName = decodeURIComponent(match[1] || match[2]);
                }
            }
        }
        // 4. 最终还拿不到文件名，拒绝处理
        if (!fileName) {
            console.error('无法确定文件名，拒绝打开');
            return;
        }
        const file = new File([blob], fileName, { type: blob.type });
        //debugger
        docmentObj.value = { fileName, file };
        showCreateDialog.value = false;
    }
    catch (err) {
        console.error('加载文件失败:', err);
        laodingInstance.close();
    }
}
onMounted(() => {
    initFileUrl();
});
debugger; /* PartiallyEnd: #3632/scriptSetup.vue */
const __VLS_ctx = {};
let __VLS_components;
let __VLS_directives;
// CSS variable injection 
// CSS variable injection end 
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "home" },
});
if (!__VLS_ctx.docmentObj?.fileName) {
    __VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
        ...{ class: "top-operation-bar" },
    });
    const __VLS_0 = {}.ElButton;
    /** @type {[typeof __VLS_components.ElButton, typeof __VLS_components.elButton, typeof __VLS_components.ElButton, typeof __VLS_components.elButton, ]} */ ;
    // @ts-ignore
    const __VLS_1 = __VLS_asFunctionalComponent(__VLS_0, new __VLS_0({
        ...{ 'onClick': {} },
        type: "primary",
    }));
    const __VLS_2 = __VLS_1({
        ...{ 'onClick': {} },
        type: "primary",
    }, ...__VLS_functionalComponentArgsRest(__VLS_1));
    let __VLS_4;
    let __VLS_5;
    let __VLS_6;
    const __VLS_7 = {
        onClick: (...[$event]) => {
            if (!(!__VLS_ctx.docmentObj?.fileName))
                return;
            __VLS_ctx.showCreateDialog = true;
        }
    };
    __VLS_3.slots.default;
    var __VLS_3;
}
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "editor-content" },
});
if (__VLS_ctx.docmentObj?.fileName) {
    /** @type {[typeof DocumentHandler, ]} */ ;
    // @ts-ignore
    const __VLS_8 = __VLS_asFunctionalComponent(DocumentHandler, new DocumentHandler({
        ...{ 'onOnlyofficeCursor': {} },
        ...{ style: {} },
        file: (__VLS_ctx.docmentObj),
        ref: "documentHandler",
    }));
    const __VLS_9 = __VLS_8({
        ...{ 'onOnlyofficeCursor': {} },
        ...{ style: {} },
        file: (__VLS_ctx.docmentObj),
        ref: "documentHandler",
    }, ...__VLS_functionalComponentArgsRest(__VLS_8));
    let __VLS_11;
    let __VLS_12;
    let __VLS_13;
    const __VLS_14 = {
        onOnlyofficeCursor: (__VLS_ctx.onOnlyofficeCursor)
    };
    /** @type {typeof __VLS_ctx.documentHandler} */ ;
    var __VLS_15 = {};
    var __VLS_10;
}
else {
    __VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
        ...{ class: "main-content" },
    });
    __VLS_asFunctionalElement(__VLS_intrinsicElements.h1, __VLS_intrinsicElements.h1)({});
    __VLS_asFunctionalElement(__VLS_intrinsicElements.p, __VLS_intrinsicElements.p)({});
}
const __VLS_17 = {}.ElDialog;
/** @type {[typeof __VLS_components.ElDialog, typeof __VLS_components.elDialog, typeof __VLS_components.ElDialog, typeof __VLS_components.elDialog, ]} */ ;
// @ts-ignore
const __VLS_18 = __VLS_asFunctionalComponent(__VLS_17, new __VLS_17({
    modelValue: (__VLS_ctx.showCreateDialog),
    title: "新建/打开文件",
    width: "450px",
    center: true,
}));
const __VLS_19 = __VLS_18({
    modelValue: (__VLS_ctx.showCreateDialog),
    title: "新建/打开文件",
    width: "450px",
    center: true,
}, ...__VLS_functionalComponentArgsRest(__VLS_18));
__VLS_20.slots.default;
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    id: "panel-createnew",
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "header" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "thumb-list" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ onClick: (...[$event]) => {
            __VLS_ctx.onCreateNew('.docx');
        } },
    ...{ class: "thumb-wrap" },
    template: "WORD",
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "thumb" },
    ...{ style: {} },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "title" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ onClick: (...[$event]) => {
            __VLS_ctx.onCreateNew('.xlsx');
        } },
    ...{ class: "thumb-wrap" },
    template: "EXCEL",
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "thumb" },
    ...{ style: {} },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "title" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ onClick: (...[$event]) => {
            __VLS_ctx.onCreateNew('.pptx');
        } },
    ...{ class: "thumb-wrap" },
    template: "PPT",
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "thumb" },
    ...{ style: {} },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "title" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "header" },
});
__VLS_asFunctionalElement(__VLS_intrinsicElements.div, __VLS_intrinsicElements.div)({
    ...{ class: "open-container" },
});
const __VLS_21 = {}.ElButton;
/** @type {[typeof __VLS_components.ElButton, typeof __VLS_components.elButton, typeof __VLS_components.ElButton, typeof __VLS_components.elButton, ]} */ ;
// @ts-ignore
const __VLS_22 = __VLS_asFunctionalComponent(__VLS_21, new __VLS_21({
    ...{ 'onClick': {} },
    type: "info",
    size: "large",
    icon: (__VLS_ctx.FolderOpened),
    plain: true,
}));
const __VLS_23 = __VLS_22({
    ...{ 'onClick': {} },
    type: "info",
    size: "large",
    icon: (__VLS_ctx.FolderOpened),
    plain: true,
}, ...__VLS_functionalComponentArgsRest(__VLS_22));
let __VLS_25;
let __VLS_26;
let __VLS_27;
const __VLS_28 = {
    onClick: (__VLS_ctx.onOpenDocument)
};
__VLS_24.slots.default;
var __VLS_24;
var __VLS_20;
/** @type {__VLS_StyleScopedClasses['home']} */ ;
/** @type {__VLS_StyleScopedClasses['top-operation-bar']} */ ;
/** @type {__VLS_StyleScopedClasses['editor-content']} */ ;
/** @type {__VLS_StyleScopedClasses['main-content']} */ ;
/** @type {__VLS_StyleScopedClasses['header']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb-list']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb-wrap']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb']} */ ;
/** @type {__VLS_StyleScopedClasses['title']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb-wrap']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb']} */ ;
/** @type {__VLS_StyleScopedClasses['title']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb-wrap']} */ ;
/** @type {__VLS_StyleScopedClasses['thumb']} */ ;
/** @type {__VLS_StyleScopedClasses['title']} */ ;
/** @type {__VLS_StyleScopedClasses['header']} */ ;
/** @type {__VLS_StyleScopedClasses['open-container']} */ ;
// @ts-ignore
var __VLS_16 = __VLS_15;
var __VLS_dollars;
const __VLS_self = (await import('vue')).defineComponent({
    setup() {
        return {
            FolderOpened: FolderOpened,
            DocumentHandler: DocumentHandler,
            showCreateDialog: showCreateDialog,
            documentHandler: documentHandler,
            docmentObj: docmentObj,
            onOnlyofficeCursor: onOnlyofficeCursor,
            onCreateNew: onCreateNew,
            onOpenDocument: onOpenDocument,
        };
    },
});
export default (await import('vue')).defineComponent({
    setup() {
        return {};
    },
});
; /* PartiallyEnd: #4569/main.vue */
//# sourceMappingURL=HomeView.vue.js.map