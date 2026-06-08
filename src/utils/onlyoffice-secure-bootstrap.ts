/** OnlyOffice web-apps 在非 HTTPS 下 crypto.subtle 不可用，生成 __guid 时会抛 digest 异常 */
export function bootstrapOnlyOfficeSecureContext(): void {
  try {
    if (!localStorage.getItem('__guid')) {
      localStorage.setItem(
        '__guid',
        `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`,
      )
    }
  } catch {
    /* 隐私模式等无法写 localStorage 时忽略 */
  }

  type CryptoWithMutableSubtle = { subtle?: SubtleCrypto }
  const g = globalThis as unknown as { crypto?: CryptoWithMutableSubtle }
  if (!g.crypto) g.crypto = {}
  if (!g.crypto.subtle) {
    g.crypto.subtle = {
      digest(_algo: AlgorithmIdentifier, data: BufferSource) {
        const bytes =
          data instanceof ArrayBuffer
            ? new Uint8Array(data)
            : new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
        const out = new Uint8Array(20)
        for (let i = 0; i < bytes.length; i++) out[i % 20] ^= bytes[i]
        return Promise.resolve(out.buffer)
      },
    } as SubtleCrypto
  }
}
