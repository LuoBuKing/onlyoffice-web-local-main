/// <reference types="vite/client" />

interface OnlyOfficeDocEditor {
  createConnector(): {
    callCommand(func: () => unknown, callback?: (result: unknown) => void): void
  }
  sendCommand(payload: { command: string; data?: unknown }): void
  destroyEditor(): void
}

interface OnlyOfficeDocsAPI {
  DocEditor: new (placeholder: string, config: Record<string, unknown>) => OnlyOfficeDocEditor
}

interface Window {
  DocsAPI?: OnlyOfficeDocsAPI
}

interface ImportMetaEnv {
  VITE_API_URL: string
  VITE_APP_TITLE: string
  VITE_PORT: number
  // 其他环境变量声明...
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
