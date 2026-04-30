// ==========================================
// アカウント移行ツール
// 使い方: GASエディタで関数を選んで▶実行するだけ
// ==========================================

// ① まずこれを実行して現状確認
function checkOwnership() {
  var folderId = getConfig().DRIVE_FOLDER_ID;
  var folder = DriveApp.getFolderById(folderId);
  var myEmail = Session.getActiveUser().getEmail();

  console.log('=== 現在の実行アカウント ===');
  console.log('メール: ' + myEmail);
  console.log('フォルダ名: ' + folder.getName());
  console.log('フォルダオーナー: ' + folder.getOwner().getEmail());

  var fileCount = 0;
  var notMineCount = 0;
  countFiles(folder, myEmail, function(owned) {
    fileCount++;
    if (!owned) notMineCount++;
  });

  console.log('総ファイル数: ' + fileCount);
  console.log('オーナーが自分でないファイル: ' + notMineCount);
  console.log(notMineCount === 0 ? '✅ 全ファイル移行済み' : '⚠️ まだ移行が必要なファイルがあります');
}

function countFiles(folder, email, callback) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    try { callback(f.getOwner().getEmail() === email); } catch(e) { callback(false); }
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) { countFiles(subs.next(), email, callback); }
}

// ② オーナー移行を実行（新アカウントのメールを入力してから実行）
function transferOwnershipToNewAccount() {
  var NEW_ACCOUNT_EMAIL = 'woodbasegroup@gmail.com';

  if (!NEW_ACCOUNT_EMAIL) {
    console.log('❌ NEW_ACCOUNT_EMAIL を入力してから実行してください');
    return;
  }

  var folderId = getConfig().DRIVE_FOLDER_ID;
  var folder = DriveApp.getFolderById(folderId);

  console.log('移行開始: ' + folder.getName() + ' → ' + NEW_ACCOUNT_EMAIL);

  var result = { success: 0, skip: 0, error: 0 };
  transferFolder(folder, NEW_ACCOUNT_EMAIL, result);

  console.log('=== 完了 ===');
  console.log('成功: ' + result.success + ' 件');
  console.log('スキップ: ' + result.skip + ' 件');
  console.log('エラー: ' + result.error + ' 件');
  console.log('次にcheckOwnership()を実行して確認してください');
}

function transferFolder(folder, newOwner, result) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      if (file.getOwner().getEmail() !== newOwner) {
        file.setOwner(newOwner);
        result.success++;
        console.log('✅ ' + file.getName());
      } else {
        result.skip++;
      }
    } catch(e) {
      result.error++;
      console.log('⚠️ スキップ: ' + file.getName() + ' (' + e.message + ')');
    }
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    transferFolder(subs.next(), newOwner, result);
  }
}
