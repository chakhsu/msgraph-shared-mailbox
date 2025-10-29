const mimeTypes = {
  // 文档类型
  pdf: 'application/pdf',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',

  // 图像类型
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tiff: 'image/tiff',
  tif: 'image/tiff',

  // 文本类型
  txt: 'text/plain',
  html: 'text/html',
  htm: 'text/html',
  css: 'text/css',
  csv: 'text/csv',
  xml: 'text/xml',
  json: 'application/json',

  // 压缩文件类型
  zip: 'application/zip',
  rar: 'application/x-rar-compressed',
  '7z': 'application/x-7z-compressed',
  tar: 'application/x-tar',
  gz: 'application/gzip',
  bz2: 'application/x-bzip2',

  // 音频类型
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  ogg: 'audio/ogg',
  flac: 'audio/flac',
  aac: 'audio/aac',

  // 视频类型
  mp4: 'video/mp4',
  avi: 'video/x-msvideo',
  mov: 'video/quicktime',
  wmv: 'video/x-ms-wmv',
  flv: 'video/x-flv',
  webm: 'video/webm',
  mkv: 'video/x-matroska',

  // 其他常见类型
  exe: 'application/x-msdownload',
  dll: 'application/x-msdownload',
  apk: 'application/vnd.android.package-archive',
  dmg: 'application/x-apple-diskimage',
  iso: 'application/x-iso9660-image',
  jar: 'application/java-archive',
  msi: 'application/x-msi',
  bin: 'application/octet-stream'
}

export const getContentTypeFromFileName = (fileName: string) => {
  const extension = fileName.split('.').pop()?.toLowerCase()
  if (!extension) {
    return 'application/octet-stream'
  }
  if (extension in mimeTypes) {
    return mimeTypes[extension as keyof typeof mimeTypes]
  } else {
    return 'application/octet-stream'
  }
}
