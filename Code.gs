/**
 * Listener Dashboard - GAS API (Code.gs)
 *
 * セキュリティ方針:
 *  - 生年月日・入社日はGAS内でのみ使用し、レスポンスに含めない
 *  - 年齢の代わりに「区分A（44歳以下）」「区分B（45歳以上）」のみ返す
 *  - new→normal切替も入社日からGASが動的判定（Sheetsは書き換えない）
 *  - 全エンドポイントでPIN認証必須（"public"/"warmup"は読み取り専用の公開データのみ）
 */

// ── スプレッドシートID（デプロイ前に設定） ──────────────
const SPREADSHEET_ID = "1rLZECadsWjB4cP63imyOjfZli27isA_30lYEeQnkx7c";

// ── シート名 ─────────────────────────────────────────────
const SHEET = {
  PLAYERS:   "Players",
  LISTENERS: "Listeners",
  RETIRED:   "Retired",
  SETTINGS:  "Settings",
  CHANGELOG: "ChangeLog",  // 変更履歴
};

// ── ブルートフォース対策の設定 ─────────────────────────────
const MAX_FAILED_PIN_ATTEMPTS = 10;      // 連続失敗この回数でロック
const PIN_LOCKOUT_MS          = 15 * 60 * 1000; // ロック時間（15分）

// ── レスポンスヘルパー ────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondError(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "error", message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Settingsシート key-value ヘルパー ─────────────────────
// key列(A)・value列(B)の単純なkey-valueストアとしてSettingsシートを読み書きする
function getSettingRow_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return { row: i + 1, value: data[i][1] };
  }
  return { row: -1, value: undefined };
}
function setSettingValue_(sheet, key, value) {
  const found = getSettingRow_(sheet, key);
  if (found.row > 0) {
    sheet.getRange(found.row, 2).setValue(value);
  } else {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(key);
    sheet.getRange(lastRow + 1, 2).setValue(value);
  }
}

// ── 排他制御（同時保存によるデータ消失を防ぐ） ─────────────
// savePlayers等の「全行クリア→書き直し」を複数リクエストが同時に行うと
// 後勝ちで片方の変更が丸ごと消えるため、書き込み系アクションはロックで直列化する
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(30 * 1000); // 30秒待って取れなければ諦める
  if (!acquired) {
    throw new Error("他の保存処理が進行中のため、少し待ってから再試行してください。");
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ── PIN照合（カウンタに影響しない、純粋な一致確認） ─────────
// 既にログイン済みの管理者が行う通常操作（savePlayers等）の認証ゲートに使う。
// ここでの失敗はブルートフォース試行ではなく「クライアントが古いPINを保持している」
// 等の可能性が高いため、失敗カウンタは増やさない（誤ロックアウト防止）。
function checkPinOnly(pin) {
  if (!pin) return false;
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET.SETTINGS);
    if (!sheet) return false;
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === "adminPass") {
        return String(data[i][1]) === String(pin);
      }
    }
    return false;
  } catch(e) {
    Logger.log("checkPinOnly error: " + e.message);
    return false;
  }
}

// ── PIN照合（ログイン試行専用・ブルートフォース対策あり） ────
// PinScreenからの明示的なログイン試行（action==="verifyPin"）でのみ使用する。
// 連続失敗が閾値を超えると一定時間ロックする。
function verifyPinWithLockout(pin) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET.SETTINGS);
    if (!sheet) return { ok: false, locked: false };

    const lockUntil = Number(getSettingRow_(sheet, "pinLockUntil").value || 0);
    if (lockUntil && Date.now() < lockUntil) {
      return { ok: false, locked: true, lockUntil: lockUntil };
    }

    const ok = checkPinOnly(pin);
    if (ok) {
      // 成功したら失敗カウンタをリセット
      setSettingValue_(sheet, "failedPinAttempts", 0);
      setSettingValue_(sheet, "pinLockUntil", "");
      return { ok: true, locked: false };
    }

    const failed = Number(getSettingRow_(sheet, "failedPinAttempts").value || 0) + 1;
    setSettingValue_(sheet, "failedPinAttempts", failed);
    if (failed >= MAX_FAILED_PIN_ATTEMPTS) {
      setSettingValue_(sheet, "pinLockUntil", Date.now() + PIN_LOCKOUT_MS);
    }
    return { ok: false, locked: false };
  } catch(e) {
    Logger.log("verifyPinWithLockout error: " + e.message);
    return { ok: false, locked: false };
  }
}

// ── 日付パース ────────────────────────────────────────────
function parseDate(val) {
  if (!val) return null;
  try {
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    // Sheetsのシリアル値（数値）対応
    if (typeof val === "number") {
      return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    const str = String(val).replace(/\//g, "-").trim();
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  } catch(e) {
    return null;
  }
}

// ── 現在の年齢を計算 ─────────────────────────────────────
// 生年月日から現在の年齢（整数）を返す。未設定の場合は null
function calcAge(dobVal) {
  const dob = parseDate(dobVal);
  if (!dob) return null;
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
  return (typeof age === "number" && !isNaN(age)) ? age : null;
}

// ── 次の期開始日を返す ────────────────────────────────────
// 「ある日付」の直後（その日を含む）に来る 1/1 or 7/1 を返す
function nextPeriodStart(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1; // 1-12
  if (m <= 6) {
    return new Date(y, 6, 1);     // 当年7月1日
  } else {
    return new Date(y + 1, 0, 1); // 翌年1月1日
  }
}

// ── 中途入社担当制限判定 ──────────────────────────────────────
// 2026-06-01以降に入社した中途入社者が入社から3年未満か判定
// 戻り値: true = 制限中（特定リスナーのみ担当可）
function isCareerRestricted(joinType, joinDateVal) {
  // 中途入社のみ担当制限対象（高卒・大卒新卒は対象外）
  if (joinType !== "career") return false;

  const joinDate = parseDate(joinDateVal);
  if (!joinDate) return false;

  // 2026-06-01以降の入社のみ対象
  const RESTRICT_START = new Date("2026-06-01");
  if (joinDate < RESTRICT_START) return false;

  // 入社から丸3年未経過 = 制限中
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const threeYearsLater = new Date(
    joinDate.getFullYear() + 3,
    joinDate.getMonth(),
    joinDate.getDate()
  );
  return today < threeYearsLater;
}

// ── new→normal 切替判定 ───────────────────────────────────
// 戻り値: { status: "new"/"normal"/"maternity", statusSoon: true/false }
//
// ルール:
//   maternity → Sheetsの値をそのまま保持（産休・育休は手動管理）
//   入社日あり → 入社日から3年経過かどうかで毎回自動判定（Sheets値を無視）
//   入社日なし → Sheetsの値をそのまま使用
//
// これによりSheetsのstatus列に誤りがあっても入社日から自動修正される
function calcStatus(currentStatus, joinDateVal) {
  // 産休・育休は入社日に関わらずそのまま保持
  if (currentStatus === "maternity") return { status: "maternity", statusSoon: false };

  const joinDate = parseDate(joinDateVal);
  // 入社日未設定 → Sheetsの値をそのまま使用
  if (!joinDate) return { status: currentStatus || "normal", statusSoon: false };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 丸3年経過日
  const threeYearsLater = new Date(
    joinDate.getFullYear() + 3,
    joinDate.getMonth(),
    joinDate.getDate()
  );

  if (threeYearsLater > today) {
    // 3年未経過 → status="new"（Sheets値に関わらず）
    // 次の期または次の次の期に切替予定かチェック（statusSoon）
    const switchDate     = nextPeriodStart(threeYearsLater);
    const nextPeriod     = nextPeriodStart(today);
    const nextNextPeriod = nextPeriodStart(nextPeriod);
    const statusSoon     = switchDate <= nextNextPeriod;
    return { status: "new", statusSoon };
  }

  // 3年経過済み → 切替日を求める
  const switchDate = nextPeriodStart(threeYearsLater);
  if (switchDate <= today) {
    // 切替日を過ぎている → "normal"
    return { status: "normal", statusSoon: false };
  } else {
    // 3年経過したが切替日未到達 → まだ"new"、次期で切替
    return { status: "new", statusSoon: true };
  }
}

// ── 45歳区分判定 ──────────────────────────────────────────
// 戻り値: { ageGroup: "A" or "B", ageSoon: true/false }
//   ageGroup: 現在の区分
//   ageSoon:  次の期または次の次の期に区分Bへ切替予定（2期先まで予告）
function calcAgeGroup(dobVal) {
  const dob = parseDate(dobVal);
  if (!dob) return { ageGroup:"A", ageSoon:false };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 45歳になる日
  const turns45 = new Date(
    dob.getFullYear() + 45,
    dob.getMonth(),
    dob.getDate()
  );

  // 切替予定日（45歳誕生日の次の期開始日）
  const switchDate = nextPeriodStart(turns45);

  // 既に切替済み（switchDate <= today）
  if (switchDate <= today) {
    return { ageGroup:"B", ageSoon:false };
  }

  // 切替前（switchDateがまだ未来）
  // → ageGroup = "A"、ageSoon = 次の期または次の次の期に切替するか
  const nextPeriod     = nextPeriodStart(today);
  const nextNextPeriod = nextPeriodStart(nextPeriod); // 2期先の開始日
  const ageSoon        = switchDate <= nextNextPeriod; // 2期以内に切替

  return { ageGroup:"A", ageSoon };
}

// ── GET: データ取得 ───────────────────────────────────────
// Players列構成:
//   A=氏名, B=生年月日, C=id, D=lid, E=status,
//   F=board, G=joinType, H=joinDate, I=returnDate,
//   J=prevLid, K=lastLidChangedAt, L=approvedMismatch
function doGet(e) {
  const pin = e.parameter.pin;
  // "public" は読み取り専用アクセス（PIN不要）
  // それ以外は正規のPIN認証
  const isPublic = (pin === "public" || pin === "warmup");
  // 修正: doGetの認証確認はログイン試行ではないため、ブルートフォース対策の
  // カウンタは増やさない checkPinOnly を使う（verifyPinWithLockoutは使わない）
  if (!isPublic && !checkPinOnly(pin)) return respondError("認証エラー: PINが正しくありません");
  // warmupリクエストは接続確認のみ（データは返さない）
  if (pin === "warmup") return respond({ warmup: true });

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Players ──────────────────────────────────────────
    const pSheet = ss.getSheetByName(SHEET.PLAYERS);
    if (!pSheet) return respondError("Playersシートが見つかりません");
    const pData  = pSheet.getDataRange().getValues();
    const players = [];

    for (let i = 1; i < pData.length; i++) {
      const row = pData[i];
      if (!row[0]) continue; // 氏名が空ならスキップ

      const name       = row[0];  // A: 氏名
      const dob        = row[1];  // B: 生年月日（レスポンスに含めない）
      const id         = row[2];  // C: id
      const lid        = row[3];  // D: lid
      const rawStatus  = row[4] || "normal"; // E: status（Sheets上の値）
      // const board = row[5]; // F: board（resolvedBoardで上書きのため未使用）
      const joinType   = row[6] || "newgrad";// G: joinType
      const joinDate   = row[7];  // H: 入社日（レスポンスに含めない）
      const returnDate = row[8];  // I: 復帰予定日（日付のみ返す）
      const prevLid    = row[9];  // J: prevLid
      const lastChanged       = row[10]; // K: lastLidChangedAt
      const approvedMismatch  = row[11] === true || String(row[11]).toLowerCase() === "true"; // L: 承認済みミスマッチ

      // ── 動的判定（生年月日・入社日はここで使い切る） ──
      // 1. new→normal の自動切替（入社から丸3年後の次の期開始日）
      const { status: resolvedStatus, statusSoon } = calcStatus(rawStatus, joinDate);
      const careerRestricted = isCareerRestricted(joinType, joinDate);

      // 2. 45歳区分（誕生日の次の期開始日から区分B）
      const { ageGroup, ageSoon } = calcAgeGroup(dob);

      // 3. board も ageGroup で上書き（区分Bなら senior）
      const resolvedBoard = ageGroup === "B" ? "senior" : "youth";

      // 4. returnDate は日付文字列のみ（生年月日ではないため返してOK）
      const returnDateStr = returnDate
        ? (parseDate(returnDate)
            ? parseDate(returnDate).toISOString().split("T")[0]
            : "")
        : "";

      // id未設定行への注意: 本来はgenId()で必ずidが付与されるため発生しないはずだが、
      // シート直接編集等でC列が空になった場合、行番号ベースのidは行の追加/削除で
      // 変動し既存の編集操作と食い違うリスクがある。空のまま放置せず気づけるようにログする。
      if (!id) {
        Logger.log("Players row " + (i + 1) + " ('" + name + "') に id が設定されていません。行番号ベースの仮IDを使用します。");
      }

      players.push({
        id:               id || ("p" + i),
        name:             name,
        ageGroup:         ageGroup,       // "A" or "B" のみ
        ageSoon:          ageSoon,        // true=次の期に区分B切替予定
        lid:              lid || "",
        status:           resolvedStatus, // 動的に判定済み
        statusSoon:       statusSoon,     // true=次の期に new→normal 切替予定
        board:            resolvedBoard,  // ageGroup から自動決定
        joinType:         joinType,
        returnDate:       returnDateStr,  // 産休・育休復帰日（個人情報ではない）
        prevLid:          prevLid   || "",
        lastLidChangedAt:  lastChanged    || "",
        approvedMismatch:  approvedMismatch,   // true=ミスマッチを承認済み
        careerRestricted:  careerRestricted,   // true=中途入社3年制限中
        age:               calcAge(dob),        // 現在の年齢（整数）※生年月日は送らない
      });
      // ※ dob, joinDate は一切レスポンスに含めない
    }

    // ── Settings から careerListeners / notes を取得 ───────
    const stSheet = ss.getSheetByName(SHEET.SETTINGS);
    let careerListenerIds = ["l1","l5","l6","l10","l12","l13","l14"]; // デフォルト値
    let adminNotes = "";
    if (stSheet) {
      const stData = stSheet.getDataRange().getValues();
      for (let i = 1; i < stData.length; i++) {
        if (String(stData[i][0]) === "careerListeners") {
          careerListenerIds = String(stData[i][1]).split(",").map(s => s.trim()).filter(s => s);
        }
        if (String(stData[i][0]) === "notes") {
          adminNotes = String(stData[i][1] || "");
        }
      }
    }

    // ── Listeners（F列=order順にソート） ─────────────────
    const lSheet    = ss.getSheetByName(SHEET.LISTENERS);
    if (!lSheet) return respondError("Listenersシートが見つかりません");
    const lData     = lSheet.getDataRange().getValues();
    const listeners = [];
    for (let i = 1; i < lData.length; i++) {
      const row = lData[i];
      if (!row[0]) continue;
      listeners.push({
        id:    row[0],
        name:  row[1],
        board: row[2],
        color: row[3],
        bg:    row[4],
        order: row[5] !== "" && row[5] !== undefined ? Number(row[5]) : i,
      });
    }
    // order昇順にソート
    listeners.sort((a, b) => a.order - b.order);

    // ── Retired ──────────────────────────────────────────
    const rSheet  = ss.getSheetByName(SHEET.RETIRED);
    if (!rSheet) return respondError("Retiredシートが見つかりません");
    const rData   = rSheet.getDataRange().getValues();
    const retired = [];
    for (let i = 1; i < rData.length; i++) {
      const row = rData[i];
      if (!row[0]) continue;
      retired.push({
        id:           row[0],
        name:         row[1],
        retiredYear:  row[2],
        retiredMonth: row[3],
        lid:          row[4],
      });
    }

    return respond({ players, listeners, retired, careerListeners: careerListenerIds, notes: adminNotes });

  } catch(err) {
    return respondError("データ取得エラー: " + err.message);
  }
}

// ── POST: データ保存 ──────────────────────────────────────
function doPost(e) {
  let body;
  try {
    if (!e.postData || !e.postData.contents) return respondError("リクエストデータがありません");
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return respondError("リクエスト形式エラー: " + err.message);
  }
  if (!body || typeof body !== "object") return respondError("リクエストが不正です");

  const action = body.action;

  // ── ログイン試行（PinScreenからの明示的な認証）── ────────
  // 修正: ここだけブルートフォース対策付きの経路を通す。
  // 他のアクション（savePlayers等）は「既にログイン済み」前提の通常確認のみ行い、
  // 失敗してもロックアウトのカウンタは増やさない（多重タブ等での誤ロック防止）。
  if (action === "verifyPin") {
    const result = verifyPinWithLockout(body.pin);
    if (result.locked) {
      const minutes = Math.max(1, Math.ceil((result.lockUntil - Date.now()) / 60000));
      return respondError(`PINの試行回数が多いため、しばらく（約${minutes}分）ロックしています。`);
    }
    if (!result.ok) return respondError("認証エラー: PINが正しくありません");
    return respond({ valid: true });
  }

  // ── それ以外のアクションは通常のPIN一致確認のみ ───────────
  if (!checkPinOnly(body.pin)) return respondError("認証エラー: PINが正しくありません");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  try {
    // 修正: 書き込み系アクションが同時実行されるとシート全体クリア→書き直しの
    // 競合で片方の変更が消えるため、ロックで直列化する
    return withLock_(() => {

      // ── Players保存 ────────────────────────────────────────
      // フロントからは既存プレイヤーの dob・joinDate を受け取らない
      // → Sheetsの既存 dob・joinDate を保持したまま、それ以外を更新
      // 新規追加プレイヤー（Sheetsに未登録=id不明）はクライアントが入力した
      // dob・joinDateをそのまま採用する（修正: 以前は常に空文字で保存されていた）
      if (action === "savePlayers") {
        const sheet   = ss.getSheetByName(SHEET.PLAYERS);
        if (!sheet) return respondError("Playersシートが見つかりません");
        const players = body.players || [];

        // ── 既存データ読み込み ─────────────────────────────────
        const existing = sheet.getDataRange().getValues();
        const dobMap   = {};   // id → dob（既存プレイヤーのみ）
        const joinMap  = {};   // id → joinDate（既存プレイヤーのみ）
        for (let i = 1; i < existing.length; i++) {
          const id = String(existing[i][2]);
          if (!id) continue;
          dobMap[id]  = existing[i][1];   // B列: 生年月日
          joinMap[id] = existing[i][7];   // H列: 入社日
        }

        // ── 安全チェック：生年月日の大幅減少を検出して中断 ─────
        // Sheetsのフィルター+ソートによる列ズレでdobMapが壊れる場合を検出
        const existingDobCount = Object.values(dobMap).filter(v => v && String(v).trim() !== "").length;
        const matchedDobCount  = players.filter(p => dobMap[p.id] && String(dobMap[p.id]).trim() !== "").length;
        if (existingDobCount >= 5 && matchedDobCount < existingDobCount * 0.85) {
          return respondError(
            `生年月日の大幅減少を検出し保存を中止しました。` +
            `（既存: ${existingDobCount}件 → 今回: ${matchedDobCount}件）` +
            `Sheetsでフィルター中に列ソートをした可能性があります。列B（生年月日）と列C（ID）の並びを確認してください。`
          );
        }

        // ── 書き込む行データを先に全て組み立てる ─────────────────
        // 修正: 以前はシートをclearContentしてから rows を組み立てていたため、
        // map中に例外が起きるとシートが空のまま残るリスクがあった。
        // 例外が起きうる処理を先に完了させ、シートへの書き込みは最後にまとめて行う。
        const rows = players.map(p => {
          const oldLid     = existing.find(r => String(r[2]) === p.id)?.[3];
          const lidChanged = oldLid && String(oldLid) !== String(p.lid);
          const approved   = lidChanged ? false : (p.approvedMismatch || false);
          const isExisting = Object.prototype.hasOwnProperty.call(dobMap, p.id);
          // 既存プレイヤー: Sheets側の値を保護してそのまま維持
          // 新規プレイヤー（isExisting=false）: クライアントが入力したdob/joinDateを採用
          const dob      = isExisting ? (dobMap[p.id]  || "") : (p.dob      || "");
          const joinDate = isExisting ? (joinMap[p.id] || "") : (p.joinDate || "");
          return [
            p.name,                      // A: 氏名
            dob,                         // B: 生年月日（既存は保護、新規は入力値を保存）
            p.id,                        // C: id
            p.lid,                       // D: lid
            p.status,                    // E: status
            p.board,                     // F: board
            p.joinType || "newgrad",     // G: joinType
            joinDate,                    // H: 入社日（既存は保護、新規は入力値を保存）
            p.returnDate       || "",    // I: returnDate
            p.prevLid          || "",    // J: prevLid
            p.lastLidChangedAt || "",    // K: lastLidChangedAt
            approved,                    // L: approvedMismatch
          ];
        });

        // ── DobBackupシートに生年月日をバックアップ ─────────────
        // 定期バックアップとして id→dob の対応を保存しておく
        // （Sheetsが壊れた際の復元に使用）
        try {
          let backupSheet = ss.getSheetByName("DobBackup");
          if (!backupSheet) {
            backupSheet = ss.insertSheet("DobBackup");
            backupSheet.getRange(1, 1, 1, 2).setValues([["id", "dob"]]);
          }
          // 既存のバックアップをクリアして最新を書き込み
          const backupLastRow = backupSheet.getLastRow();
          if (backupLastRow > 1) backupSheet.getRange(2, 1, backupLastRow - 1, 2).clearContent();
          // rows組み立て後の最新dob（新規プレイヤー分も含む）をバックアップ対象にする
          const backupRows = rows
            .filter(r => r[1] && String(r[1]).trim() !== "")
            .map(r => [r[2], r[1]]); // [id, dob]
          if (backupRows.length > 0) {
            backupSheet.getRange(2, 1, backupRows.length, 2).setValues(backupRows);
          }
        } catch(e) {
          // バックアップ失敗は無視（メイン処理を止めない）
          Logger.log("DobBackup error: " + e.message);
        }

        // ── ここまでで例外が起きなければ、シートをクリアして書き直す ─────
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
        }
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, 12).setValues(rows);
        }
        return respond({ saved: rows.length });
      }

      // ── Retired保存 ────────────────────────────────────────
      if (action === "saveRetired") {
        const sheet   = ss.getSheetByName(SHEET.RETIRED);
        if (!sheet) return respondError("Retiredシートが見つかりません");
        const retired = body.retired || [];

        // 修正: クリア前にrowsを組み立てておく（例外時の空振りクリアを避ける）
        const rows = retired.map(r => [
          r.id, r.name, r.retiredYear, r.retiredMonth, r.lid,
        ]);

        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
        }
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, 5).setValues(rows);
        }
        return respond({ saved: rows.length });
      }

      // ── リスナー並び順保存 ─────────────────────────────────
      if (action === "saveListenerOrder") {
        const order = body.order || []; // [{ id:"l1", order:0 }, ...]
        const sheet = ss.getSheetByName(SHEET.LISTENERS);
        if (!sheet) return respondError("Listenersシートが見つかりません");
        const data  = sheet.getDataRange().getValues();
        // F列にorder値を書き込み
        for (let i = 1; i < data.length; i++) {
          const id  = data[i][0];
          const hit = order.find(o => o.id === id);
          if (hit !== undefined) {
            sheet.getRange(i + 1, 6).setValue(hit.order);
          }
        }
        return respond({ saved: order.length });
      }

      // ── パスワード変更 ──────────────────────────────────────
      if (action === "changePass") {
        const newPass = String(body.newPass || "");
        // 修正: クライアント側ログイン画面(PinScreen)は数字4桁固定の専用キーパッドのため、
        // サーバー側でも同じ制約を強制する（以前はlength<4のみで、5桁以上や英字を許してしまい
        // 直接API経由でPINを設定するとログイン自体が不可能になるロックアウトの恐れがあった）
        if (!/^\d{4}$/.test(newPass)) {
          return respondError("PINは数字4桁で設定してください");
        }
        const sheet = ss.getSheetByName(SHEET.SETTINGS);
        if (!sheet) return respondError("Settingsシートが見つかりません");
        const data  = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === "adminPass") {
            sheet.getRange(i + 1, 2).setValue(newPass);
            return respond({ changed: true });
          }
        }
        return respondError("Settings シートに adminPass が見つかりません");
      }

      // ── 変更履歴保存 ────────────────────────────────────────
      if (action === "saveChangeLog") {
        const logs = body.logs || [];
        if (logs.length === 0) return respond({ saved: 0 });
        let target = ss.getSheetByName(SHEET.CHANGELOG);
        if (!target) {
          // シートがなければ作成してヘッダーを設定
          target = ss.insertSheet(SHEET.CHANGELOG);
          if (!target) return respondError("ChangeLogシートの作成に失敗しました");
          target.getRange(1,1,1,5).setValues([["日時","プレイヤー名","変更内容","変更前","変更後"]]);
        }
        const rows = logs.map(l => [
          l.datetime  || "",
          l.playerName|| "",
          l.changeType|| "",
          l.before    || "",
          l.after     || "",
        ]);
        const last = target.getLastRow();
        target.getRange(last + 1, 1, rows.length, 5).setValues(rows);
        return respond({ saved: rows.length });
      }

      // ── 備考保存 ────────────────────────────────────────────
      if (action === "saveNotes") {
        const notes = String(body.notes || "");
        const sheet = ss.getSheetByName(SHEET.SETTINGS);
        if (!sheet) return respondError("Settingsシートが見つかりません");
        const data = sheet.getDataRange().getValues();
        // 既存の notes 行を更新
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === "notes") {
            sheet.getRange(i + 1, 2).setValue(notes);
            return respond({ saved: true });
          }
        }
        // なければ新規追加
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1).setValue("notes");
        sheet.getRange(lastRow + 1, 2).setValue(notes);
        return respond({ saved: true });
      }

      // ── 変更履歴取得（最新100件、新しい順） ─────────────────
      if (action === "getChangelog") {
        const clSheet = ss.getSheetByName(SHEET.CHANGELOG);
        if (!clSheet) return respond({ logs: [] });
        const clData = clSheet.getDataRange().getValues();
        const logs = [];
        for (let i = 1; i < clData.length; i++) {
          if (!clData[i][0]) continue; // 空行スキップ
          logs.push({
            datetime:   String(clData[i][0] || ""),
            playerName: String(clData[i][1] || ""),
            changeType: String(clData[i][2] || ""),
            before:     String(clData[i][3] || ""),
            after:      String(clData[i][4] || ""),
          });
        }
        // 最新100件を新しい順で返す
        return respond({ logs: logs.slice(-100).reverse() });
      }

      return respondError("不明なaction: " + action);
    });

  } catch(err) {
    return respondError("保存エラー: " + err.message);
  }
}

// ── keepAlive: コールドスタート防止（案D）───────────────────────
// GASトリガーで5分ごとに呼び出すことでコールドスタートを排除する
// 【設定手順】
// 1. GASエディタを開く
// 2. 左メニューの「時計アイコン（トリガー）」をクリック
// 3. 右下「トリガーを追加」をクリック
// 4. 設定：
//    - 実行する関数: keepAlive
//    - イベントのソース: 時間主導型
//    - 時間ベースのトリガーのタイプ: 分ベースのタイマー
//    - 時間の間隔: 5分おき
// 5. 保存（Googleアカウント認証が求められたら許可）
function keepAlive() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET.PLAYERS);
    if (sheet) sheet.getRange("A1").getValue();
    Logger.log("keepAlive OK: " + new Date().toISOString());
  } catch(e) {
    Logger.log("keepAlive Error: " + e.message);
  }
}

// ── DobBackupから生年月日を復元（手動実行用） ────────────────────
// 使い方：GASエディタからこの関数を手動で実行する
// 生年月日が消えた場合に、DobBackupシートの値をPlayersシートのB列に復元する
function restoreDobFromBackup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // バックアップシートを確認
  const backupSheet = ss.getSheetByName("DobBackup");
  if (!backupSheet) {
    Logger.log("DobBackupシートが見つかりません。savePlayers実行後に自動作成されます。");
    return;
  }

  // バックアップからid→dobマップを作成
  const backupData = backupSheet.getDataRange().getValues();
  const backupMap = {};
  for (let i = 1; i < backupData.length; i++) {
    const id  = String(backupData[i][0]);
    const dob = backupData[i][1];
    if (id && dob) backupMap[id] = dob;
  }
  Logger.log("バックアップ件数: " + Object.keys(backupMap).length);

  // Playersシートに復元
  const playersSheet = ss.getSheetByName("Players");
  if (!playersSheet) {
    Logger.log("Playersシートが見つかりません。");
    return;
  }
  const playersData = playersSheet.getDataRange().getValues();
  let restored = 0;
  for (let i = 1; i < playersData.length; i++) {
    const id  = String(playersData[i][2]); // C列: id
    const dob = playersData[i][1];         // B列: 現在の生年月日
    if (id && backupMap[id] && (!dob || String(dob).trim() === "")) {
      // 生年月日が空の行にバックアップから復元
      playersSheet.getRange(i + 1, 2).setValue(backupMap[id]);
      restored++;
      Logger.log(`復元: row ${i+1}, id=${id}, dob=${backupMap[id]}`);
    }
  }
  Logger.log(`復元完了: ${restored}件`);
}
