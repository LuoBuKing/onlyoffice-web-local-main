<template>
  <div class="home">
    <div class="top-operation-bar" v-if="!route.query.url && !docmentObj?.fileName">
      <el-button type="primary" @click="showCreateDialog = true">新建/打开文件</el-button>
    </div>
    <div
      class="editor-content"
      v-loading="pageLoading"
      element-loading-text="文档加载中..."
      element-loading-background="rgba(255, 255, 255, 0.85)"
    >
      <DocumentHandler
        v-if="docmentObj?.fileName"
        style="height: 100%; width: 100%"
        :file="docmentObj"
        :readonly="editorReadonly"
        ref="documentHandler"
        @onlyoffice-cursor="onOnlyofficeCursor"
      />
    </div>

    <el-dialog v-model="showCreateDialog" title="新建/打开文件" width="450px" center>
      <div id="panel-createnew">
        <div class="header">新建</div>
        <div class="thumb-list">
          <div class="thumb-wrap" template="WORD" @click="onCreateNew('.docx')">
            <div class="thumb" style="background-image: url('./img/doc-formats/docx.png')"></div>
            <div class="title">文档</div>
          </div>
          <div class="thumb-wrap" template="EXCEL" @click="onCreateNew('.xlsx')">
            <div class="thumb" style="background-image: url('./img/doc-formats/xlsx.png')"></div>
            <div class="title">表格</div>
          </div>
          <div class="thumb-wrap" template="PPT" @click="onCreateNew('.pptx')">
            <div class="thumb" style="background-image: url('./img/doc-formats/pptx.png')"></div>
            <div class="title">演示文稿</div>
          </div>
        </div>
        <div class="header">打开</div>
        <div class="open-container">
          <el-button type="info" size="large" :icon="FolderOpened" @click="onOpenDocument" plain>
            打开本地文件
          </el-button>
        </div>
      </div>
    </el-dialog>
  </div>
</template>

<script lang="ts" setup>
import { FolderOpened } from '@element-plus/icons-vue'
import { computed, onMounted, ref, watch } from 'vue'
import { DocmentType } from '@/utils/util'
import DocumentHandler from '../components/DocumentHandler.vue'
import { useRoute } from 'vue-router'

const showCreateDialog = ref(false)
const documentHandler = ref<InstanceType<typeof DocumentHandler> | null>(null)
const docmentObj = ref<DocmentType | null>(null)
const pageLoading = ref(false)

function onOnlyofficeCursor(payload: {
  event: string
  source?: string
  doc?: Record<string, unknown>
}) {
  console.log('[OnlyOffice 光标探针]', payload.event, payload.source ?? '', payload.doc ?? '(no doc)')
}

const onCreateNew = (ext: string) => {
  pageLoading.value = false
  docmentObj.value = {
    fileName: '新建文档' + ext,
    file: null,
  }
  showCreateDialog.value = false
}

const onOpenDocument = () => {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.docx,.xlsx,.pptx,.doc,.xls,.ppt'

  input.onchange = (event) => {
    const file = (event.target as HTMLInputElement).files?.[0]
    if (file) {
      pageLoading.value = false
      showCreateDialog.value = false
      docmentObj.value = {
        fileName: file.name,
        file: file,
      }
    }
  }

  input.click()
}

const route = useRoute()

const editorReadonly = computed(() => {
  const ro = route.query.readonly
  return ro === '1' || ro === 'true'
})

async function initFileUrl() {
  const url = route.query.url as string | undefined
  const filenameParam = route.query.filename as string | undefined
  if (!url) {
    pageLoading.value = false
    return
  }

  pageLoading.value = true
  try {
    const res = await fetch(url)
    if (!res.ok) throw new Error('文件请求失败')

    const blob = await res.blob()
    let fileName = ''

    if (filenameParam) {
      fileName = filenameParam
    }
    if (!fileName) {
      const match = decodeURIComponent(url).match(/\/([^\/?#]+)$/)
      if (match && match[1].includes('.')) {
        fileName = match[1]
      }
    }
    if (!fileName) {
      const disposition = res.headers.get('Content-Disposition')
      if (disposition) {
        const match = disposition.match(/filename\*=UTF-8''(.+)|filename="?([^"]+)"?/)
        if (match) {
          fileName = decodeURIComponent(match[1] || match[2])
        }
      }
    }
    if (!fileName) {
      console.error('无法确定文件名，拒绝打开')
      return
    }

    const file = new File([blob], fileName, { type: blob.type })
    docmentObj.value = { fileName, file }
    showCreateDialog.value = false
  } catch (err) {
    console.error('加载文件失败:', err)
  } finally {
    pageLoading.value = false
  }
}

onMounted(() => {
  pageLoading.value = !!route.query.url
  void initFileUrl()
})

watch(
  () => route.query.url,
  (url) => {
    if (url) void initFileUrl()
  },
)
</script>

<style lang="less" scoped>
.home {
  display: flex;
  flex-direction: column;
  height: 100vh;
  background-color: #f5f5f5;
}

.top-operation-bar {
  background-color: white;
  padding: 12px 20px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.1);
  z-index: 10;
}

.editor-content {
  flex: 1;
  min-height: 0;
  position: relative;
}

#panel-createnew {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1);
  padding: 20px;
  .header {
    font-size: 18px;
    padding: 0 0 0 25px;
    white-space: nowrap;
    margin-top: 20px;
    margin-bottom: 20px;
  }

  .thumb-list {
    display: flex;
    justify-content: space-around;

    .thumb-wrap {
      display: inline-block;
      text-align: center;
      width: auto;
      cursor: pointer;
      vertical-align: top;
      border-radius: 4px;

      .thumb {
        width: 96px;
        height: 96px;
        background-repeat: no-repeat;
        background-position: center;
        margin: 12px 12px 0px 12px;
        background-size: contain;
      }

      .title {
        width: 104px;
        font-size: 14px;
        line-height: 14px;
        height: 28px;
        margin: 8px 8px 12px 8px;
        word-break: break-word;
        word-wrap: break-word;
        overflow: hidden;
        text-overflow: ellipsis;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        line-clamp: 2;
        -webkit-box-orient: vertical;
      }

      &:hover {
        background-color: #e0e0e0;
      }

      &:active {
        color: rgba(0, 0, 0, 0.8);
        background-color: #cbcbcb;
      }
    }
  }
}
.open-container {
  text-align: center;
  padding-bottom: 25px;
  margin-top: 20px;
}
</style>
