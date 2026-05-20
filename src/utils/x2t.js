/**
 * X2T 工具类 - 负责文档转换功能
 */
class X2TConverter {
    x2tModule = null;
    isReady = false;
    initPromise = null;
    hasScriptLoaded = false;
    // 支持的文件类型映射
    DOCUMENT_TYPE_MAP = {
        docx: 'word',
        doc: 'word',
        odt: 'word',
        rtf: 'word',
        txt: 'word',
        xlsx: 'cell',
        xls: 'cell',
        ods: 'cell',
        csv: 'cell',
        pptx: 'slide',
        ppt: 'slide',
        odp: 'slide',
    };
    WORKING_DIRS = [
        '/working',
        '/working/media',
        '/working/fonts',
        '/working/themes',
    ];
    SCRIPT_PATH = './wasm/x2t/x2t.js';
    INIT_TIMEOUT = 20000;
    /**
     * 加载 X2T 脚本文件
     */
    async loadScript() {
        if (this.hasScriptLoaded)
            return;
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = this.SCRIPT_PATH;
            script.onload = () => {
                this.hasScriptLoaded = true;
                console.log('X2T WASM script loaded successfully');
                resolve();
            };
            script.onerror = (error) => {
                const errorMsg = 'Failed to load X2T WASM script';
                console.error(errorMsg, error);
                reject(new Error(errorMsg));
            };
            document.head.appendChild(script);
        });
    }
    /**
     * 初始化 X2T 模块
     */
    async initialize() {
        if (this.isReady && this.x2tModule) {
            return this.x2tModule;
        }
        // 防止重复初始化
        if (this.initPromise) {
            return this.initPromise;
        }
        this.initPromise = this.doInitialize();
        return this.initPromise;
    }
    async doInitialize() {
        try {
            // 确保脚本已加载
            await this.loadScript();
            return new Promise((resolve, reject) => {
                const x2t = window.Module;
                if (!x2t) {
                    reject(new Error('X2T module not found after script loading'));
                    return;
                }
                // 设置超时处理
                const timeoutId = setTimeout(() => {
                    if (!this.isReady) {
                        reject(new Error(`X2T initialization timeout after ${this.INIT_TIMEOUT}ms`));
                    }
                }, this.INIT_TIMEOUT);
                x2t.onRuntimeInitialized = () => {
                    try {
                        clearTimeout(timeoutId);
                        this.createWorkingDirectories(x2t);
                        this.x2tModule = x2t;
                        this.isReady = true;
                        console.log('X2T module initialized successfully');
                        resolve(x2t);
                    }
                    catch (error) {
                        reject(error);
                    }
                };
            });
        }
        catch (error) {
            this.initPromise = null; // 重置以允许重试
            throw error;
        }
    }
    /**
     * 创建工作目录
     */
    createWorkingDirectories(x2t) {
        this.WORKING_DIRS.forEach((dir) => {
            try {
                x2t.FS.mkdir(dir);
            }
            catch (error) {
                // 目录可能已存在，忽略错误
                console.warn(`Directory ${dir} may already exist:`, error);
            }
        });
    }
    /**
     * 获取文档类型
     */
    getDocumentType(extension) {
        const docType = this.DOCUMENT_TYPE_MAP[extension.toLowerCase()];
        if (!docType) {
            throw new Error(`Unsupported file format: ${extension}`);
        }
        return docType;
    }
    /**
     * 清理文件名
     */
    sanitizeFileName(input) {
        if (typeof input !== 'string' || !input.trim()) {
            return 'file.bin';
        }
        const parts = input.split('.');
        const ext = parts.pop() || 'bin';
        const name = parts.join('.');
        const illegalChars = /[\/\?<>\\:\*\|"]/g;
        const controlChars = /[\x00-\x1f\x80-\x9f]/g;
        const reservedPattern = /^\.+$/;
        const unsafeChars = /[&'%!"{}[\]]/g;
        let sanitized = name
            .replace(illegalChars, '')
            .replace(controlChars, '')
            .replace(reservedPattern, '')
            .replace(unsafeChars, '');
        sanitized = sanitized.trim() || 'file';
        return `${sanitized.slice(0, 200)}.${ext}`; // 限制长度
    }
    /**
     * 执行文档转换
     */
    executeConversion(paramsPath) {
        if (!this.x2tModule) {
            throw new Error('X2T module not initialized');
        }
        const result = this.x2tModule.ccall('main1', 'number', ['string'], [paramsPath]);
        if (result !== 0) {
            throw new Error(`Conversion failed with code: ${result}`);
        }
    }
    /**
     * 创建转换参数XML
     */
    createConversionParams(fromPath, toPath, additionalParams = '') {
        return `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>${fromPath}</m_sFileFrom>
  <m_sThemeDir>/working/themes</m_sThemeDir>
  <m_sFileTo>${toPath}</m_sFileTo>
  <m_bIsNoBase64>false</m_bIsNoBase64>
  ${additionalParams}
</TaskQueueDataConvert>`;
    }
    /**
     * 读取媒体文件
     */
    readMediaFiles() {
        if (!this.x2tModule)
            return {};
        const media = {};
        try {
            const files = this.x2tModule.FS.readdir('/working/media/');
            files
                .filter((file) => file !== '.' && file !== '..')
                .forEach((file) => {
                try {
                    const fileData = this.x2tModule.FS.readFile(`/working/media/${file}`, {
                        encoding: 'binary',
                    });
                    const blob = new Blob([fileData]);
                    const mediaUrl = URL.createObjectURL(blob);
                    media[`media/${file}`] = mediaUrl;
                }
                catch (error) {
                    console.warn(`Failed to read media file ${file}:`, error);
                }
            });
        }
        catch (error) {
            console.warn('Failed to read media directory:', error);
        }
        return media;
    }
    /**
     * 将文档转换为 bin 格式
     */
    async convertDocument(file) {
        await this.initialize();
        const fileName = file.name;
        const fileExt = fileName.split('.').pop()?.toLowerCase() || '';
        const documentType = this.getDocumentType(fileExt);
        try {
            // 读取文件内容
            const arrayBuffer = await file.arrayBuffer();
            const data = new Uint8Array(arrayBuffer);
            // 生成安全的文件名
            const sanitizedName = this.sanitizeFileName(fileName);
            const inputPath = `/working/${sanitizedName}`;
            const outputPath = `${inputPath}.bin`;
            // 写入文件到虚拟文件系统
            this.x2tModule.FS.writeFile(inputPath, data);
            // 创建转换参数
            const params = this.createConversionParams(inputPath, outputPath);
            this.x2tModule.FS.writeFile('/working/params.xml', params);
            // 执行转换
            this.executeConversion('/working/params.xml');
            // 读取转换结果
            const result = this.x2tModule.FS.readFile(outputPath);
            const media = this.readMediaFiles();
            return {
                fileName: sanitizedName,
                type: documentType,
                bin: result,
                media,
            };
        }
        catch (error) {
            throw new Error(`Document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    }
    /**
     * 将 bin 格式转换为指定格式并下载
     */
    async convertBinToDocument(bin, originalFileName, targetExt = 'DOCX') {
        await this.initialize();
        const sanitizedBase = this.sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, '');
        const binFileName = `${sanitizedBase}.bin`;
        const outputFileName = `${sanitizedBase}.${targetExt.toLowerCase()}`;
        try {
            this.x2tModule.FS.writeFile(`/working/${binFileName}`, bin);
            let additionalParams = '';
            if (targetExt === 'PDF') {
                additionalParams = '<m_sFontDir>/working/fonts/</m_sFontDir>';
            }
            const params = this.createConversionParams(`/working/${binFileName}`, `/working/${outputFileName}`, additionalParams);
            this.x2tModule.FS.writeFile('/working/params.xml', params);
            this.executeConversion('/working/params.xml');
            const result = this.x2tModule.FS.readFile(`/working/${outputFileName}`);
            return {
                fileName: outputFileName,
                data: result,
            };
        }
        catch (error) {
            throw new Error(`Bin to document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    }
    async convertBinToDocumentAndDownload(bin, originalFileName, targetExt = 'DOCX') {
        const result = await this.convertBinToDocument(bin, originalFileName, targetExt);
        this.saveWithFileSystemAPI(result.data, result.fileName);
        return result;
    }
    /**
     * 下载文件
     */
    downloadFile(data, fileName) {
        const blob = new Blob([data]);
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        // 清理资源
        setTimeout(() => {
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, 100);
    }
    /**
     * 根据文件扩展名获取 MIME 类型
     */
    getMimeTypeFromExtension(extension) {
        const mimeMap = {
            // 文档类型
            docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            doc: 'application/msword',
            odt: 'application/vnd.oasis.opendocument.text',
            rtf: 'application/rtf',
            txt: 'text/plain',
            pdf: 'application/pdf',
            // 表格类型
            xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            xls: 'application/vnd.ms-excel',
            ods: 'application/vnd.oasis.opendocument.spreadsheet',
            csv: 'text/csv',
            // 演示文稿类型
            pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            ppt: 'application/vnd.ms-powerpoint',
            odp: 'application/vnd.oasis.opendocument.presentation',
            // 图片类型
            png: 'image/png',
            jpg: 'image/jpeg',
            jpeg: 'image/jpeg',
            gif: 'image/gif',
            bmp: 'image/bmp',
            webp: 'image/webp',
            svg: 'image/svg+xml',
        };
        return mimeMap[extension.toLowerCase()] || 'application/octet-stream';
    }
    /**
     * 获取文件类型描述
     */
    getFileDescription(extension) {
        const descriptionMap = {
            docx: 'Word Document',
            doc: 'Word 97-2003 Document',
            odt: 'OpenDocument Text',
            pdf: 'PDF Document',
            xlsx: 'Excel Workbook',
            xls: 'Excel 97-2003 Workbook',
            ods: 'OpenDocument Spreadsheet',
            pptx: 'PowerPoint Presentation',
            ppt: 'PowerPoint 97-2003 Presentation',
            odp: 'OpenDocument Presentation',
            txt: 'Text Document',
            rtf: 'Rich Text Format',
            csv: 'CSV File',
        };
        return descriptionMap[extension.toLowerCase()] || 'Document';
    }
    /**
     * 使用现代文件系统 API 保存文件
     */
    async saveWithFileSystemAPI(data, fileName, mimeType) {
        if (!window.showSaveFilePicker) {
            this.downloadFile(data, fileName);
            return;
        }
        try {
            // 获取文件扩展名并确定 MIME 类型
            const extension = fileName.split('.').pop()?.toLowerCase() || '';
            const detectedMimeType = mimeType || this.getMimeTypeFromExtension(extension);
            // 显示文件保存对话框
            const fileHandle = await window.showSaveFilePicker({
                suggestedName: fileName,
                types: [
                    {
                        description: this.getFileDescription(extension),
                        accept: {
                            [detectedMimeType]: [`.${extension}`],
                        },
                    },
                ],
            });
            // 创建可写流并写入数据
            const writable = await fileHandle.createWritable();
            await writable.write(data);
            await writable.close();
            console.log('File saved successfully:', fileName);
        }
        catch (error) {
            if (error.name === 'AbortError') {
                console.log('User cancelled the save operation');
                return;
            }
            throw error;
        }
    }
    /**
     * 销毁实例，清理资源
     */
    destroy() {
        this.x2tModule = null;
        this.isReady = false;
        this.initPromise = null;
        console.log('X2T converter destroyed');
    }
}
// 单例实例
const x2tConverter = new X2TConverter();
// 导出的公共 API
export const initX2TScript = () => x2tConverter.loadScript();
export const initX2T = () => x2tConverter.initialize();
export const convertDocument = (file) => x2tConverter.convertDocument(file);
export const convertBinToDocument = (bin, fileName, targetExt) => x2tConverter.convertBinToDocument(bin, fileName, targetExt);
export const convertBinToDocumentAndDownload = (bin, fileName, targetExt) => x2tConverter.convertBinToDocumentAndDownload(bin, fileName, targetExt);
// 文件类型常量
export const oAscFileType = {
    UNKNOWN: 0,
    PDF: 513,
    PDFA: 521,
    DJVU: 515,
    XPS: 516,
    DOCX: 65,
    DOC: 66,
    ODT: 67,
    RTF: 68,
    TXT: 69,
    HTML: 70,
    MHT: 71,
    EPUB: 72,
    FB2: 73,
    MOBI: 74,
    DOCM: 75,
    DOTX: 76,
    DOTM: 77,
    FODT: 78,
    OTT: 79,
    DOC_FLAT: 80,
    DOCX_FLAT: 81,
    HTML_IN_CONTAINER: 82,
    DOCX_PACKAGE: 84,
    OFORM: 85,
    DOCXF: 86,
    DOCY: 4097,
    CANVAS_WORD: 8193,
    JSON: 2056,
    XLSX: 257,
    XLS: 258,
    ODS: 259,
    CSV: 260,
    XLSM: 261,
    XLTX: 262,
    XLTM: 263,
    XLSB: 264,
    FODS: 265,
    OTS: 266,
    XLSX_FLAT: 267,
    XLSX_PACKAGE: 268,
    XLSY: 4098,
    PPTX: 129,
    PPT: 130,
    ODP: 131,
    PPSX: 132,
    PPTM: 133,
    PPSM: 134,
    POTX: 135,
    POTM: 136,
    FODP: 137,
    OTP: 138,
    PPTX_PACKAGE: 139,
    IMG: 1024,
    JPG: 1025,
    TIFF: 1026,
    TGA: 1027,
    GIF: 1028,
    PNG: 1029,
    EMF: 1030,
    WMF: 1031,
    BMP: 1032,
    CR2: 1033,
    PCX: 1034,
    RAS: 1035,
    PSD: 1036,
    ICO: 1037,
};
export const c_oAscFileType2 = Object.fromEntries(Object.entries(oAscFileType).map(([key, value]) => [value, key]));
//# sourceMappingURL=x2t.js.map