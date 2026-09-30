// ===== 設定 =====
// ボードごとに、Apps Script を「ウェブアプリ」としてデプロイした URL（…/exec）を貼り付ける。
// URL が空欄のボードはデモモード（サンプルデータ・保存されない）で動きます。
//
// 開き方： 学生スタッフ      https://tnagaibio-ops.github.io/task-board/
//          実習アシスタント  https://tnagaibio-ops.github.io/task-board/?b=assist
// （QR は qr.html でボードを選んで作成すると、この形の URL になります）
window.TASKBOARD_CONFIG = {
  POLL_SECONDS: 20,          // 自動更新の間隔（秒）
  DEFAULT_BOARD: 'staff',    // ?b= を付けずに開いたときのボード
  BOARDS: {
    staff: {
      label: '学生スタッフ',
      GAS_URL: 'https://script.google.com/macros/s/AKfycbxZr7vI8hgmD64vS6cQZlJFHYDpq8iutPSSyapGuJE6EabPM30Ovb3MdSLD7_e78Lu0/exec',
    },
    assist: {
      label: '実習アシスタント',
      GAS_URL: 'https://script.google.com/macros/s/AKfycbxdjmV2J2OKoFvCU69ClOQmggepOOVHC9ibdaUfFzW9z1SL84f0Elyh0CSlABdczzHs-w/exec', // ← 実習アシスタント用スプレッドシートのウェブアプリURLを貼る
    },
  },
};
