// ============================================================
// Attachment_API.gs
// Code.gs に追記するか、新規ファイルとして追加する
// 製品マスタ・議事録・メモ・機種・基板への添付ファイル管理
// ============================================================

/**
 * 親レコードに紐づく添付ファイル一覧を返す
 * @param {string} parentId   - 親レコードのID
 * @param {string} parentType - 'meeting'|'note'|'product'|'model'|'board'|'kb'
 */
function getAttachments(parentId, parentType) {
  let all = rows('attachments');
  all = all.filter(a => String(a['親ID']) === String(parentId));
  if (parentType) all = all.filter(a => a['親種別'] === parentType);
  return all.map(a => ({
    id:       a['添付ID'],
    name:     a['ファイル名'],
    fileId:   a['DriveファイルID'],
    url:      a['Drive URL'],
    date:     a['登録日時'],
    type:     a['親種別'],
    mimeType: a['MIMEタイプ'] || '',
    size:     a['ファイルサイズ'] || '',
  }));
}

/**
 * base64ファイルをDriveに保存して attachments シートに記録
 * 親レコードの「添付DriveURL」列も更新する
 */
function uploadAttachmentFull(parentId, parentType, fileName, base64Data, mimeType, fileSizeKB) {
  try {
    // 保存先フォルダを取得 or 作成
    const prop = PropertiesService.getScriptProperties();
    let rootFid = prop.getProperty(PROP_ATTACH);
    let rootFolder;
    if (rootFid) {
      try { rootFolder = DriveApp.getFolderById(rootFid); }
      catch(e) { rootFolder = null; }
    }
    if (!rootFolder) {
      rootFolder = DriveApp.getRootFolder().createFolder('SKH_Attachments');
      prop.setProperty(PROP_ATTACH, rootFolder.getId());
    }

    // 親種別サブフォルダ
    const subIter = rootFolder.getFoldersByName(parentType);
    const sub = subIter.hasNext() ? subIter.next() : rootFolder.createFolder(parentType);

    // ファイル保存
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, fileName);
    const file    = sub.createFile(blob);

    // attachments シートに記録
    const aid = uid();

    // シートヘッダー拡張チェック（MIMEタイプ・サイズ列がなければ追加しない、既存構造維持）
    push('attachments', [
      aid, parentId, parentType,
      fileName, file.getId(), file.getUrl(), now()
    ]);

    // 親レコードの「添付DriveURL」を更新（URLを改行区切りで追記）
    const sheetMap = {
      meeting: { sheet:'meetings', col:'添付DriveURL' },
      note:    { sheet:'notes',    col:'添付DriveURL' },
      product: { sheet:'products', col:'添付DriveURL' },
      model:   { sheet:'models',   col:'添付DriveURL' },
      board:   { sheet:'boards',   col:'添付DriveURL' },
      kb:      { sheet:'kb_articles', col:'添付DriveURL' },
    };
    const target = sheetMap[parentType];
    if (target) {
      const rec = rows(target.sheet).find(r => {
        const firstKey = Object.keys(r)[0];
        return String(r[firstKey]) === String(parentId);
      });
      if (rec) {
        const existing = rec[target.col] || '';
        const newVal   = existing
          ? existing + '\n' + file.getUrl()
          : file.getUrl();
        updateRowData(target.sheet, parentId, target.col, newVal);
      }
    }

    return {
      status:  'success',
      id:      aid,
      fileId:  file.getId(),
      url:     file.getUrl(),
      name:    fileName,
      mimeType,
      sizeKB:  fileSizeKB || Math.round(decoded.length / 1024),
    };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 添付ファイルを削除する
 * Drive上のファイルを削除し、attachments シートからも行を削除
 */
function deleteAttachment(attachId) {
  try {
    const all = rows('attachments');
    const rec = all.find(a => String(a['添付ID']) === String(attachId));
    if (!rec) return { status: 'error', message: '添付が見つかりません' };

    // Driveからファイルを削除（ゴミ箱へ）
    try {
      const file = DriveApp.getFileById(rec['DriveファイルID']);
      file.setTrashed(true);
    } catch(e) {
      // ファイルが既に削除されていてもシートからは削除する
      Logger.log('Drive削除スキップ: ' + e.message);
    }

    // attachments シートから行削除
    deleteRowData('attachments', attachId);

    // 親レコードの「添付DriveURL」からそのURLを除去
    const sheetMap = {
      meeting: { sheet:'meetings', col:'添付DriveURL' },
      note:    { sheet:'notes',    col:'添付DriveURL' },
      product: { sheet:'products', col:'添付DriveURL' },
      model:   { sheet:'models',   col:'添付DriveURL' },
      board:   { sheet:'boards',   col:'添付DriveURL' },
      kb:      { sheet:'kb_articles', col:'添付DriveURL' },
    };
    const target = sheetMap[rec['親種別']];
    if (target) {
      const parentRec = rows(target.sheet).find(r => String(Object.values(r)[0]) === String(rec['親ID']));
      if (parentRec) {
        const existing = parentRec[target.col] || '';
        const newVal   = existing
          .split('\n')
          .filter(u => u.trim() && u.trim() !== rec['Drive URL'].trim())
          .join('\n');
        updateRowData(target.sheet, rec['親ID'], target.col, newVal);
      }
    }

    return { status: 'success', message: `「${rec['ファイル名']}」を削除しました` };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 親レコードの添付ファイル件数をまとめて返す（テーブル表示用）
 * { parentId: count } の形で返す
 */
function getAttachmentCounts(parentType) {
  const all = rows('attachments').filter(a => a['親種別'] === parentType);
  const map = {};
  all.forEach(a => {
    const pid = String(a['親ID']);
    map[pid] = (map[pid] || 0) + 1;
  });
  return map;
}
