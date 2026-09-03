# ==============================================================================
# 火祭り資源と組織調査結果 - 統計分析スクリプト（v3）
# データ: data_raw/火祭の資源と組織調査結果.xlsx
# 作成日: 2026-06-27
# 出力先: data_proc/
#
# 分析項目:
#   01 協力者年齢（複数協力者を個別表示）
#   02 関係者数・観光客数トレンド（横ばい=黄色）
#   03 植物資源出現頻度（全件）
#   04 植物資源種数（祭り別）
#   05 組織の有無
#   06 祭り目的キーワード頻度
#   07 信仰キーワード頻度
#   08 近年の課題（自由記述のキーワード分類）
#   09 植物資源 × 祭り マトリクス
#   10 関係者数 vs 観光客数 散布図
#   11 規模（関係者数） vs 植物資源種数
#   12 生息地多様性 vs 植物資源種数
#   13 植物資源 × 信仰 クロス分析
#   14 植物資源 × 祭り目的 クロス分析
#   --- 生物文化多様性分析 ---
#   15 文化的关键种（代替可能性 × 出現頻度）
#   16 文化-生態嵌入度（調達方法の地域性）
#   17 伝統生態知識（TEK）の深さ（植物の選定理由）
#   18 生息地依存ネットワーク（祭り × 景観タイプ）
#   19 文化-生態脆弱性（嵌入度変化 × 課題）
# ==============================================================================

library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(ggplot2)
library(forcats)
library(ggrepel)   # install.packages("ggrepel") if needed
library(patchwork) # 図01の左右並置に使用

# 2026-09-02 更新：30シート版（新規16祭り追加）に切替
# 旧: DATA_PATH <- "data_raw/火祭の資源と組織調査結果.xlsx"  (14シート, 出力先 data_proc/)
DATA_PATH <- "data_raw/火祭の資源と組織調査結果_20260830.xlsx"
OUTPUT_DIR <- "data_proc/20260902"
dir.create(OUTPUT_DIR, showWarnings = FALSE, recursive = TRUE)

# ------------------------------------------------------------------------------
# 祭り所在府県（図の並び順のグループ化に使用）
# ------------------------------------------------------------------------------
# 出典：data_raw/近畿地方火祭り.xlsx（都道府県・市町村欄）および調査票の
#       「関係社寺名」「氏子地域範囲」欄。順序は本研究の近畿圏の定義に合わせる。
PREF_ORDER <- c("滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県")

FESTIVAL_PREF <- c(
  # 滋賀県
  "巽神社松明"               = "滋賀県",  # 近江八幡市糠塚町（奥石神社が祭礼に関与）
  "松明を次世代に送る会"     = "滋賀県",  # 近江八幡市
  "雄琴学区ヨシ松明一斉点火" = "滋賀県",  # 大津市雄琴学区
  "太郎坊宮の火祭り"         = "滋賀県",  # 東近江市小脇町
  "信楽の火祭り"             = "滋賀県",  # 甲賀市信楽町
  "勝部の火祭り"             = "滋賀県",  # 守山市勝部町
  "近江八幡左義長祭り"       = "滋賀県",  # 近江八幡市（日牟禮八幡宮）
  "八幡祭り"                 = "滋賀県",  # 近江八幡市（日牟禮八幡宮）
  "王の浜若宮神社"           = "滋賀県",  # 近江八幡市白王町王の浜
  "小田神社"                 = "滋賀県",  # 近江八幡市北里学区
  "大嶋奥津嶋神社"           = "滋賀県",  # 近江八幡市北津田町
  # 京都府
  "鞍馬の火祭"               = "京都府",  # 京都市左京区鞍馬
  "三栖の火祭"               = "京都府",  # 京都市伏見区三栖
  "大文字送り火"             = "京都府",  # 京都市左京区浄土寺
  "嵯峨のお松明式"           = "京都府",  # 京都市右京区嵯峨（清凉寺）
  "広河原松上げ"             = "京都府",  # 京都市左京区広河原
  "花背松上げ"               = "京都府",  # 京都市左京区花脊八桝町
  "雲ケ畑松上げ"             = "京都府",  # 京都市北区雲ヶ畑出谷町
  # 大阪府
  "がんがら火祭り"           = "大阪府",  # 池田市（旧池田村）
  "まんどろ火祭り"           = "大阪府",  # 箕面市萱野
  "麦わら松明"               = "大阪府",  # 箕面市北芝
  # 兵庫県
  "東光寺鬼会"               = "兵庫県",  # 加西市上万願寺（国指定重要無形民俗文化財）
  "稲引き樽引き神事"         = "兵庫県",  # 三田市賀茂（加茂神社）
  "湯村火祭り"               = "兵庫県",  # 美方郡新温泉町湯
  # 奈良県
  "往馬大社"                 = "奈良県",  # 生駒市
  "吉祥草寺茅原大とんど"     = "奈良県",  # 御所市茅原
  "ほうらんや火祭り"         = "奈良県",  # 橿原市東部六地区
  # 和歌山県
  "熊野速玉大社"             = "和歌山県",  # 新宮市（御燈祭り）
  "稲むらの火祭り"           = "和歌山県",  # 有田郡広川町
  "熊野那智"                 = "和歌山県"   # 東牟婁郡那智勝浦町
)

# ------------------------------------------------------------------------------
# 抽出調査の府県別ウェイト（事後層化）
# ------------------------------------------------------------------------------
# 【なぜ必要か】
# 30祭りの調査対象は府県ごとの抽出率が大きく異なる。素の「利用祭り数」は
# 抽出の多い府県で使われる植物を過大に、抽出の少ない府県に偏在する植物を
# 過小に見せてしまう。そこで母集団（近畿地方火祭り.xlsx で把握した156件）の
# 府県構成に合わせて事後層化ウェイト w = N_府県 / n_府県 を与える。
#
# POP_FRAME は data_raw/近畿地方火祭り.xlsx の都道府県欄の集計。
# 同ファイルで都道府県が空欄の3件（箕面市・堺市・河内長野市）は市町村から
# 大阪府に算入している。
#
# 【限界（結果の解釈時に必ず併記すること）】
#   1. 156件の母集団自体が悉皆ではない（本文でも「網羅するものではない」と明記）。
#      特に大阪府は4件と極端に少なく、実態より過小に把握されている可能性が高い。
#      → 大阪府のウェイトは小さくなりすぎている恐れがある。
#   2. 府県内の抽出は無作為ではない（有意抽出）。ウェイトが補正するのは
#      府県間の構成比のみで、府県内の選択バイアスは補正できない。
#   3. 府県あたりの標本が3〜11件と小さく、府県内の比率は粗い（3件なら0/⅓/⅔/1）。
#      → 層化ブートストラップで区間を併記する。
POP_FRAME <- c("滋賀県" = 79, "京都府" = 30, "大阪府" = 4,
               "兵庫県" = 21, "奈良県" = 12, "和歌山県" = 10)

pref_weights <- function(sample_pref) {
  n_smp <- table(factor(sample_pref, levels = PREF_ORDER))
  w <- POP_FRAME[PREF_ORDER] / as.numeric(n_smp)
  tibble(pref = factor(PREF_ORDER, levels = PREF_ORDER),
         n_sample = as.numeric(n_smp),
         n_pop    = as.numeric(POP_FRAME[PREF_ORDER]),
         w        = as.numeric(w))
}

# ------------------------------------------------------------------------------
# 0. ユーティリティ関数
# ------------------------------------------------------------------------------

# 行ラベルに対応するセル値をシートから取得
get_row_values <- function(sheet_df, label_pat) {
  col1_clean <- str_replace_all(as.character(sheet_df[[1]]), "\\s+", "")
  match_rows <- which(str_detect(col1_clean, label_pat))
  if (length(match_rows) == 0) return(NA_character_)
  vals <- as.character(sheet_df[match_rows[1], -1])
  vals <- vals[!is.na(vals) & vals != "NULL" & vals != "NA"]
  if (length(vals) == 0) return(NA_character_)
  paste(vals, collapse = " / ")
}

# 変化傾向の正規化
parse_trend <- function(x) {
  x <- str_replace_all(as.character(x), "\\s+", "")
  case_when(
    str_detect(x, "大量増加")         ~ "大幅増加",
    str_detect(x, "増加")             ~ "増加",
    str_detect(x, "ほぼ同じ|変化なし")~ "横ばい",
    str_detect(x, "大量減少")         ~ "大幅減少",
    str_detect(x, "減少")             ~ "減少",
    str_detect(x, "^なし$|^0$")       ~ "なし/0",
    TRUE                              ~ "不明"
  )
}

# 年齢セルから協力者の年齢を抽出（1セルに複数名・自由記述が混在する）
# ------------------------------------------------------------------------------
# 【2026-09-02 改訂の理由】0830版では年齢欄の記入形式が多様化したため、
# 単純な最初の数値の抜き出しでは誤りが生じていた：
#   「田島氏2004年生まれ（2026年時点21～22歳）」→ 生年 2004 を年齢として拾う
#   「個別年齢は未確認。中心メンバーは30代～70代…80歳目前」→ 30 を年齢として拾う
#   「角田氏 64歳／西村氏 46歳。」→ 1セル内の2人目 46 を落とす
# 方針：
#   1. 「年齢は未確認」と明記されたセルは協力者個人の年齢なしとして扱う
#   2. 「NN歳」形式があればそれを全て採用（生年・世代の数値を拾わない）
#      範囲併記（「21～22歳」）は上限を採用
#   3. 「歳」がない場合のみ、10〜110の範囲の裸の数値を年齢とみなす
extract_ages <- function(x) {
  s <- str_replace_all(as.character(x), "[\r\n]+", " ")
  if (is.na(s) || str_trim(s) %in% c("", "NA", "NULL")) return(numeric(0))
  if (str_detect(s, "年齢は未確認|年齢未確認")) return(numeric(0))
  with_sai <- str_match_all(s, "([0-9]{1,3})\\s*歳")[[1]]
  if (nrow(with_sai) > 0) {
    v <- as.numeric(with_sai[, 2])
  } else {
    v <- as.numeric(str_extract_all(s, "[0-9]{1,3}")[[1]])
  }
  v <- v[!is.na(v) & v >= 10 & v <= 110]
  v
}

# 資源名の正規化（表記ゆれ統一）
normalize_resource <- function(x) {
  x %>%
    str_replace_all("（.*?）|\\(.*?\\)", "") %>%
    str_replace_all("\\s+", "") %>%
    str_replace_all("稲藁|稲わら|稲ワラ|いねわら", "稲わら") %>%
    str_replace_all("ヨシ（葦）|ヨシ$|葦$", "ヨシ") %>%
    str_replace_all("タケ$|竹$|枯れ竹.*|笹$|ササ$|竹・笹|タケ・笹", "タケ/竹") %>%
    str_replace_all("ヒノキ.*", "ヒノキ") %>%
    str_replace_all("アカマツ|赤松.*|マツ$|マツ・.*", "マツ") %>%
    str_replace_all("菜種.*|ナタネ.*", "菜種殻") %>%
    str_replace_all("スギ.*|杉.*", "スギ") %>%
    str_replace_all("食材各種", "食材（各種）") %>%
    str_replace_all("小麦の麦わら|麦わら", "麦わら") %>%
    str_replace_all("柴：.*|柴木.*", "柴") %>%
    str_replace_all("フジツル|ツツラフジ", "フジ・ツル類") %>%
    str_replace_all("木の芯棒.*|杉丸太.*", "杉丸太") %>%
    str_trim() %>%
    canon_resource()
}

# 2026-09-02 追加：30シート版で新たに現れた表記ゆれの最終統合
# （上の正規表現チェーンで拾い切れないものを辞書で一括統合）
CANON_RESOURCE <- c(
  # タケ類
  "真タケ/竹" = "タケ/竹", "タケ類" = "タケ/竹", "ナヨタケ/竹" = "タケ/竹",
  "竹・タケ/竹" = "タケ/竹",
  # 稲わら・わら製品
  "藁縄" = "稲わら", "わら縄" = "稲わら", "稲わら縄" = "稲わら",
  "もちわら" = "稲わら", "稲わら・縄・ムシロ・藁製品" = "稲わら",
  # 麦わら（「小麦ワラ」はカタカナ表記のため上のルールで拾えない）
  "小麦ワラ" = "麦わら", "小麦わら" = "麦わら",
  # ススキ・カヤ
  "ススギ" = "ススキ", "すすぎ" = "ススキ", "茅" = "カヤ",
  # フジ・ツル類
  "藤" = "フジ・ツル類", "藤蔓" = "フジ・ツル類", "フジ" = "フジ・ツル類",
  # その他の同義語
  "椎木" = "シイ", "苧殻" = "麻ガラ", "麻ロープ" = "麻縄",
  # 肥松＝ジン（樹脂化したマツ材）。樹種名（アカマツ・クロマツ）を問わず統合
  "肥松" = "肥松/ジン", "クロマツの肥松" = "肥松/ジン", "マツのジン" = "肥松/ジン",
  # マツの葉
  "マツの枝・マツの葉" = "マツの葉", "マツの松葉" = "マツの葉",
  # 非植物資材（下流で除外する）
  "綿タオル" = "非植物資材", "古布・タオル" = "非植物資材",
  "アルミ製升・灯油" = "非植物資材"
)

canon_resource <- function(x) {
  hit <- x %in% names(CANON_RESOURCE)
  x[hit] <- unname(CANON_RESOURCE[x[hit]])
  x
}

# ------------------------------------------------------------------------------
# 資源名 → 植物分類群（タクソン）への集約
# ------------------------------------------------------------------------------
# 【2026-09-02 追加の理由】
# 0830版の調査票は資源名を「樹種＋部位・形態」で細分して記録している
# （例：ヒノキの丸太／ヒノキの薪／ヒノキの葉／ヒノキのかんなくず）。
# 「植物資源の種類数」を数えるには部位を落として分類群に集約する必要がある。
# ※ 部位別の細かさが必要な分析では resource_norm（記録どおり）を使うこと。
TAXON_RULES <- list(
  c("非植物資材",        "アルミ|灯油|タオル|古布|綿"),
  c("松竹梅（飾り）",    "松竹梅"),
  c("マツ",              "アカマツ|クロマツ|マツ|肥松|ジン|赤松"),
  c("ヒノキ",            "ヒノキ|檜"),
  c("スギ",              "スギ|杉"),
  c("タケ/ササ類",       "タケ|竹|ササ|笹"),
  c("イネ（稲わら）",    "稲ワラ|稲わら|稲藁|藁縄|わら縄|もちわら|もち米|赤米|稲穂|^稲"),
  c("ムギ（麦わら）",    "小麦|麦わら|麦ワラ"),
  c("ナタネ（菜種ガラ）","菜種|ナタネ"),
  c("ヨシ",              "ヨシ|葦"),
  c("ススキ",            "ススキ|ススギ|すすぎ"),
  c("カヤ",              "カヤ|^茅"),
  c("アサ（麻）",        "麻ガラ|麻縄|麻ロープ|苧殻"),
  c("フジ類",            "フジ|藤|ツツラフジ"),
  c("ツツジ類",          "ツツジ|コバノミツバツツジ"),
  c("クロモジ",          "クロモジ"),
  c("ソヨゴ",            "ソヨゴ"),
  c("ハンノキ",          "ハンノキ"),
  c("シイ",              "シイ|椎"),
  c("ヌルデ",            "ヌルデ"),
  c("シキミ",            "シキミ"),
  c("サカキ",            "サカキ|榊"),
  c("ツバキ",            "ツバキ"),
  c("ヒオウギ",          "ヒオウギ"),
  c("クリ",              "クリ|栗"),
  c("吉祥草",            "吉祥草"),
  c("雑木・広葉樹類",    "雑木|広葉樹|^柴"),
  c("食材（各種）",      "食材"),
  c("樹種不明の木材",    "木の丸太|木樽|^薪$")
)

normalize_taxon <- function(x) {
  v <- str_replace_all(as.character(x), "（.*?）|\\(.*?\\)", "")
  v <- str_replace_all(v, "\\s+", "")
  out <- rep(NA_character_, length(v))
  for (r in TAXON_RULES) {
    hit <- is.na(out) & !is.na(v) & str_detect(v, r[2])
    out[hit] <- r[1]
  }
  out[is.na(out) & !is.na(v)] <- v[is.na(out) & !is.na(v)]  # 未分類はそのまま残す
  out
}

# ------------------------------------------------------------------------------
# 1. 全シート読み込み
# ------------------------------------------------------------------------------

sheet_ids <- excel_sheets(DATA_PATH)          # 読み込み用（原文のまま）
sheets    <- str_trim(sheet_ids)                # 表示・結合キー用（前後空白を除去）
# 旧版では「小田神社 」の末尾空白により scale_df との結合が外れ、
# n_resources が NA になっていた。ここで統一する。
cat("シート数:", length(sheets), "\n")

survey_list <- lapply(sheet_ids, function(sh) {
  df <- read_excel(DATA_PATH, sheet = sh, col_names = FALSE)
  list(
    festival         = str_trim(sh),
    age_raw          = get_row_values(df, "^年齢$"),
    belief           = get_row_values(df, "^信仰$"),
    purpose          = get_row_values(df, "^祭り目的$"),
    participants_r7  = get_row_values(df, "祭り関係者数"),
    participants_chg = get_row_values(df, "祭り関係者数変化"),
    tourists_r7      = get_row_values(df, "祭り観光客数"),
    tourists_chg     = get_row_values(df, "祭り観光客数変化"),
    preservation     = get_row_values(df, "^保存会$"),
    ujiko_org        = get_row_values(df, "^氏子組織$"),
    admin_org        = get_row_values(df, "^関係する行政"),
    company          = get_row_values(df, "^企業・財団$"),
    school           = get_row_values(df, "^学校$"),
    core_generation  = get_row_values(df, "^火祭り中心世代$"),
    challenges       = get_row_values(df, "^近年の課題"),
    social_meaning   = get_row_values(df, "^火祭りの社会意義$")
  )
})

survey_df <- bind_rows(lapply(survey_list, as.data.frame, stringsAsFactors = FALSE)) %>%
  mutate(
    participants_trend = parse_trend(participants_chg),
    tourists_trend     = parse_trend(tourists_chg)
  )

# ------------------------------------------------------------------------------
# 1b. 協力者年齢 — 長形式（複数協力者を個別に展開）
# ------------------------------------------------------------------------------

age_long <- lapply(sheet_ids, function(sh) {
  df <- read_excel(DATA_PATH, sheet = sh, col_names = FALSE)
  age_row <- df[str_replace_all(as.character(df[[1]]), "\\s+", "") == "年齢", ]
  if (nrow(age_row) == 0) return(NULL)
  vals <- as.character(age_row[1, -1])
  vals <- vals[!is.na(vals) & vals != "NA"]
  ages <- unlist(lapply(vals, extract_ages))
  if (length(ages) == 0) return(NULL)
  data.frame(festival = str_trim(sh), age = ages)
}) %>% bind_rows()

# ------------------------------------------------------------------------------
# 1c. 植物資源 + 生息地 — 長形式
# ------------------------------------------------------------------------------

resource_raw <- lapply(sheet_ids, function(sh) {
  df <- read_excel(DATA_PATH, sheet = sh, col_names = FALSE)
  col1_clean <- str_replace_all(as.character(df[[1]]), "\\s+", "")

  resource_header <- which(col1_clean == "資源名")
  group_row <- which(str_detect(col1_clean, "グループにおける植物資源"))

  if (length(resource_header) == 0 || length(group_row) == 0) return(NULL)
  start <- resource_header[1] + 1
  end   <- group_row[1] - 1
  if (start > end) return(NULL)

  # ヘッダー行からカラム位置を特定
  hdr <- as.character(df[resource_header[1], ])
  get_col <- function(pat) { i <- which(str_detect(hdr, pat)); if (length(i)) i[1] else NA_integer_ }
  landscape_col <- get_col("景観")
  use_col       <- get_col("利用方法")
  method_col    <- get_col("調達方法")
  subst_col     <- get_col("代替")
  reason_col    <- get_col("選定理由")
  daily_col     <- get_col("日常利用")

  rows <- df[start:end, ]
  resource_names <- str_replace_all(as.character(rows[[1]]), "\\s+", "")
  valid <- resource_names != "" & !is.na(resource_names) & resource_names != "NA"

  get_col_vals <- function(col_idx) {
    if (!is.na(col_idx)) as.character(rows[[col_idx]]) else rep(NA_character_, nrow(rows))
  }

  data.frame(
    festival      = str_trim(sh),
    resource_raw  = as.character(rows[[1]])[valid],
    landscape_raw = get_col_vals(landscape_col)[valid],
    use_raw       = get_col_vals(use_col)[valid],
    method_raw    = get_col_vals(method_col)[valid],
    subst_raw     = get_col_vals(subst_col)[valid],
    reason_raw    = get_col_vals(reason_col)[valid],
    daily_raw     = get_col_vals(daily_col)[valid],
    stringsAsFactors = FALSE
  )
}) %>% bind_rows()

# ------------------------------------------------------------------------------
# 代替可能性スコア：
#   3 = 代用できない（文化的に不可欠）
#   2 = 代用品あるが使わない（部分的不可欠）
#   1 = 代用できる（代替可能）
# 調達方法の地域嵌入度スコア：
#   3 = 共同体が自ら採取・栽培（最高嵌入）
#   2 = 地域農家・知人からもらう（中嵌入）
#   1 = 市場・業者から購入（最低嵌入）
# TEKタイプ（植物の選定理由）：
#   eco  = 生態的特性（燃焼性・強度・形状など）
#   aes  = 審美的・感覚的（色・香り・見た目）
#   trad = 伝統・慣習（昔から・伝統・慣習）
#   prag = 実用的可及性（多い・手に入りやすい）
#   symb = 象徴的・宗教的（縁起・神聖・奉納）
# 利用方法の分類：
#   燃焼材  = 松明の主要燃料・着火材
#   構造材  = 芯棒・骨組み・軸・充填材
#   化粧材  = 外部化粧・外装・外側を覆う
#   装飾材  = 飾り・装飾・縁起物・笠
#   結束材  = くくる・縛る・縄・巻く（ただし主用途が化粧・装飾でないもの）
#   食材    = ダシ・食材
#   ※ 複数カテゴリーをパイプ区切りで記録
# ------------------------------------------------------------------------------

code_use <- function(x) {
  if (is.na(x) || str_trim(x) %in% c("", "NA", "NULL", "None")) return(NA_character_)
  cats <- character(0)
  if (str_detect(x, "燃料|燃焼材|燃焼物|火床で燃やす|着火|燃え|主要材料.*燃|燃える"))
    cats <- c(cats, "燃焼材")
  if (str_detect(x, "芯棒|骨組み|心棒|中心軸|充填|軸|枠|内部|中に入れ|中身|本体"))
    cats <- c(cats, "構造材")
  if (str_detect(x, "化粧|外側|外装|外部|仕上が|外面"))
    cats <- c(cats, "化粧材")
  if (str_detect(x, "飾り|装飾|縁起|鱗|笠|頭|先頭部|山車|ダシ|食材|奉納"))
    cats <- c(cats, "装飾材")
  if (str_detect(x, "くくる|縛|結束|縄|巻く|締め") &&
      !any(c("化粧材","装飾材") %in% cats))
    cats <- c(cats, "結束材")
  if (str_detect(x, "ダシ|食材|食べ"))
    cats <- c(cats, "食材")
  if (length(cats) == 0) cats <- "その他"
  paste(cats, collapse = "|")
}

# 日常利用スコア：
#   1 = 日常的に使う
#   2 = ほとんどない
#   3 = 全くない
# ------------------------------------------------------------------------------

code_daily <- function(x) {
  if (is.na(x) || str_trim(x) %in% c("", "NA", "NULL", "None")) return(NA_integer_)
  x_clean <- str_replace_all(as.character(x), "\\s+|\n.*", "")
  case_when(
    str_detect(x_clean, "^１|^1|日常的に使う|日常利用あり") ~ 1L,
    str_detect(x_clean, "^２|^2|ほとんどない")              ~ 2L,
    str_detect(x_clean, "^３|^3|全くない")                  ~ 3L,
    TRUE ~ NA_integer_
  )
}

# ------------------------------------------------------------------------------

code_substitutability <- function(x) {
  x <- str_replace_all(as.character(x), "\\s+", "")
  case_when(
    str_detect(x, "^３|^3|代用できない|代用.*困難") ~ 3L,
    str_detect(x, "^２|^2|代用品あるが|しない方がいい|限定的") ~ 2L,
    str_detect(x, "^１|^1|代用できる|代替可能") ~ 1L,
    TRUE ~ NA_integer_
  )
}

code_embeddedness <- function(x) {
  x_clean <- str_replace_all(as.character(x), "\n", " ")
  case_when(
    # 最高嵌入：共同体自採・自栽培・市民活動
    str_detect(x_clean,
      "市民.*刈り|住民が.*刈|住民が.*採|住民が.*探|住民が.*掘|集団活動|自分たちで刈|栽培する|栽培し|神田で|保存会が栽培|各隣組が.*栽培|隣組が.*田んぼ|手刈り|共同.*採取|境内.*切り出す|竹やぶから切|山から切り出") ~ 3L,
    # 中嵌入：地域農家・知人・神社・組合から無償
    str_detect(x_clean,
      "農家.*持ってくる|農家.*もらう|農家にお願い|地域住民が持|地域の人が|持ち主.*了解|無償.*もらう|もらう|農家から|森林組合.*もらえる|町内.*採取|地域内で取") ~ 2L,
    # 最低嵌入：購入（業者・ネット・市場）
    str_detect(x_clean,
      "購入|通販|業者|林業家から|木材屋|造園屋|花屋|竹屋|ヨシ屋|ヨシ業者|建材屋|丸山.*ヨシ屋") ~ 1L,
    TRUE ~ NA_integer_
  )
}

# ------------------------------------------------------------------------------
# 選定理由の類型（2026-09-02 新設）
# ------------------------------------------------------------------------------
# 【従来の code_tek との違い】
# 旧 TEK 5類型（eco/aes/trad/prag/symb）は 0830版の選定理由の自由記述を
# 読み直すと以下を区別できていなかった：
#   ・燃焼特性（燃えやすさ・火力・持続）と 物理/加工特性（まっすぐ・軽い・強度）
#   ・単なる入手容易性と、生業の副産物であること（里山・農林業との結合）
#   ・本来の選好ではなく、資源枯渇・コスト・技術低下による代替選択
#   ・安全・子供の参加・世代継承といった社会的機能
#   ・環境保全そのものを目的とする選択（ヨシ刈り、間伐材利用）
# そこで156件の原文から帰納的に10類型を立てた。1記録が複数類型を持つ
# （実際 156件中104件が2類型以上）。
#   burn   燃焼特性       燃えやすい・火力・持続時間・油分・煙
#   phys   物理/加工特性  まっすぐ・軽い・強度・しなやか・加工しやすい・寸法
#   sens   感覚/美的      色・香り・見た目・音・緑・清浄感・装飾性
#   avail  入手容易性     手に入りやすい・地域に多い・身近・調達が容易
#   byprod 生業副産物     農業/林業の副産物・裏作・間伐材・不要材の循環利用
#   trad   伝統/慣習      昔からの材料・伝統・由来・継承
#   symb   象徴/宗教      縁起・神聖・魔除け・奉納・豊穣の象徴・伝承
#   subst  代替/制約      本命が確保できない/高価/技術低下のための代替選択
#   social 社会的機能     子供の参加・世代継承・安全・村同士の競い合い
#   env    環境保全       水質浄化・環境保護政策・里山保全を目的とする選択
REASON_RULES <- list(
  c("burn",  "燃えやす|燃えにく|火力|火勢|火付|火つき|着火|燃焼|長持ち|燃え残|持続|油分|松脂|樹脂|煙|温度|よく燃え|最後まで燃|燃やしやす|派手に燃え|消えにく|火に強|燃え方|火よけ|燃え上が"),
  c("phys",  "まっすぐ|真っ直ぐ|直材|直線性|軽い|軽く|軽さ|重い|強度|強さ|丈夫|しなやか|しなり|柔軟|柔らか|加工しやす|扱いやす|使いやす|割り加工|太さ|太く|長さ|長く|径|折れにく|粘り|中空|細く|形状|曲がりが少な|寸法|かさ|束ねやす|締めやす|切れにく|構造材|骨組み|芯|支柱|型崩れ"),
  c("sens",  "色|綺麗|きれい|美し|香り|見た目|肌が|緑|音|白っぽ|清浄|外観|見栄え|装飾|飾り"),
  c("avail", "手に入りやす|入手|調達|身近|近くに|地域に多い|豊富|得やす|確保できる|多く生え|自生|地域内|周辺で|供給|地域から|地域産|地域で得|地域にある|容易"),
  c("byprod","副産物|裏作|二毛作|間伐|循環利用|不要材|廃棄|再利用|余った|農産物|農作物|無駄なく"),
  c("trad",  "昔から|伝統|慣習|昔の|継承|由来|古くから|歴史|従来|旧来|昔ぼ|開始時|定着"),
  c("symb",  "縁起|神聖|めでた|魔除|奉納|象徴|祈願|神事|神が宿|常若|信仰|伝承|清め|しめ縄|注連縄|御幣|紙垂|供物|五穀豊穣|豊作|神に|神前|祈るための木"),
  c("subst", "代替|代用|代わり|確保できない|不足|安価|コスト|補強|技術力|技術が弱く|現実的|貴重|模倣"),
  c("social","子供|子ども|教育|次世代|関心|対抗意識|見せ場|競う|遊び|楽しみ|安全|ささくれ|危険"),
  c("env",   "水質|浄化|環境保護|環境保全|里山保全")
)

REASON_LABELS <- c(
  burn   = "燃焼特性",
  phys   = "物理・加工特性",
  sens   = "感覚・美的",
  avail  = "入手容易性",
  byprod = "生業副産物",
  trad   = "伝統・慣習",
  symb   = "象徴・宗教",
  subst  = "代替・制約",
  social = "社会的機能",
  env    = "環境保全"
)

code_reason <- function(x) {
  s <- str_replace_all(as.character(x), "[\r\n]+", " ")
  if (is.na(s) || str_trim(s) %in% c("", "NA", "NULL")) return(NA_character_)
  if (str_detect(s, "^不明$|^理由なし|^未確認")) return(NA_character_)
  # 「田遊び」は農耕儀礼の名称であり social の「遊び」ではない（誤検出の回避）
  s <- str_replace_all(s, "田遊び", "田儀礼")
  ty <- character(0)
  for (r in REASON_RULES) if (str_detect(s, r[2])) ty <- c(ty, r[1])
  if (!length(ty)) return(NA_character_)
  paste(ty, collapse = "|")
}

code_tek <- function(x) {
  x <- str_replace_all(as.character(x), "\n", " ")
  if (is.na(x) || str_trim(x) %in% c("", "NA", "NULL")) return(NA_character_)
  types <- character(0)
  if (str_detect(x, "燃えやす|燃焼|火力|中空|軽い|強度|柔軟|火に強|油が多|まっすぐ|音がよ|バチバチ|水分|着火|長持ち|細くて|温度"))
    types <- c(types, "eco")
  if (str_detect(x, "色が|綺麗|香り|見た目|肌が|形が|緑が|清浄|美し"))
    types <- c(types, "aes")
  if (str_detect(x, "昔から|伝統|慣習|歴史|昔のやり方|昔の.*材料|慣行|由来"))
    types <- c(types, "trad")
  if (str_detect(x, "手に入りやす|多い|入手が容易|調達.*便利|確保できる|近くで|豊富|農産物"))
    types <- c(types, "prag")
  if (str_detect(x, "縁起|神聖|奉納|めでた|象徴|清め|豊作.*証|神事|神様|御幣|紙垂"))
    types <- c(types, "symb")
  if (length(types) == 0) return(NA_character_)
  paste(types, collapse = "|")
}

# ------------------------------------------------------------------------------
# 生息地（景観）の正規化
# ------------------------------------------------------------------------------
# 【2026-09-02 改訂の理由】
# 14シート版では景観欄が統制語彙（森林／湿地／水田／畑／荒地／庭）で記入されて
# いたが、30シート版の新規16シートは自由記述（例「賀茂周辺の赤松林・里山。冬に
# 枯死・樹脂化した松を探す山林利用。」）で記入されている。旧来の前方一致置換では
# 分類できないため、キーワードによる優先順位付き分類に変更した。
# ・旧6カテゴリーは維持し、新データで頻出する2類型を追加：
#     草地       … ススキ群落・草本景観など二次草原（旧「荒地」とは区別）
#     流通・生活空間 … 製材所・木材流通・購入・各家庭など景観と呼べない調達元
# ・「茅原」は地名（御所市茅原）でもあるため草地キーワードから除外している。
# ・判定結果は landscape_mapping_check.csv に原文つきで出力するので要確認。
classify_landscape <- function(x) {
  v <- str_replace_all(as.character(x), "[\r\n]+", " ")
  out <- rep(NA_character_, length(v))
  rules <- list(
    c("流通・生活空間", "^製材所|^木材加工|木材流通|樽製造|企業協力|各家庭|旅館|温泉観光施設|生活・観光空間|野外活動施設"),
    c("水田",  "水田|田んぼ|稲作|田の隅|神田"),
    c("湿地",  "ヨシ原|よし原|湿地|河川|湖岸|水辺|川沿い"),
    c("畑",    "畑|農地|圃場"),
    c("草地",  "ススキ群落|草本景観|草地|原っぱ|草原|ススキ状"),
    c("森林",  "森林|山林|里山|竹林|竹薮|竹やぶ|松林|雑木林|二次林|社叢|境内|神社の森|樹林|山地|丘陵|林縁|林道|防災林|伐採跡|林$|林。|林・|林、"),
    c("庭",    "庭園|^庭$|屋敷林"),
    c("荒地",  "荒地|空地|放棄地"),
    c("流通・生活空間", "流通|購入|業者|外部流通|市場")
  )
  for (r in rules) {
    hit <- is.na(out) & !is.na(v) & str_detect(v, r[2])
    out[hit] <- r[1]
  }
  out[!is.na(v) & str_detect(v, "^(なし|NA|NULL|不明)$")] <- NA_character_
  out
}

resource_df <- resource_raw %>%
  mutate(
    resource_norm = normalize_resource(resource_raw),
    # 生息地の正規化（classify_landscape() を参照）
    landscape_norm = classify_landscape(landscape_raw),
    # 代替可能性スコア（1〜3）
    subst_score = code_substitutability(subst_raw),
    # 調達方法の嵌入度スコア（1〜3）
    embed_score = code_embeddedness(method_raw),
    # TEKタイプ（複数）
    tek_types = mapply(code_tek, reason_raw),
    # 利用方法カテゴリー（複数可）
    use_types = mapply(code_use, use_raw),
    # 日常利用スコア（1〜3）
    daily_score = mapply(code_daily, daily_raw),
    # 植物分類群（部位・形態を落とした集約。種類数のカウントに使う）
    resource_taxon = normalize_taxon(resource_raw),
    # 選定理由の類型（複数可、パイプ区切り）
    reason_types = vapply(reason_raw, code_reason, character(1))
  ) %>%
  # 長すぎる行（注記として混入したもの）と空白・NAを除去
  filter(
    !is.na(resource_norm),
    str_trim(resource_norm) != "",
    resource_norm != "NA",
    nchar(resource_norm) <= 20,
    !str_detect(resource_norm, "[。！？]|使用不可|注意"),
    # 2026-09-02 追加：本分析は「植物資源」を対象とするため非植物資材を除外
    # （調査票側で「非植物補助資材」と注記されたアルミ製升・灯油、古布・タオル等）
    resource_norm != "非植物資材"
  )

# ==============================================================================
# 植物 × 祭り の解析単位と、府県ウェイトによる推定
# ==============================================================================
# 【解析単位】 祭り × 植物分類群（festival × taxon）
#   ・同じ祭りで同じ植物の部位が複数記録されていても（ヒノキの丸太／薪／葉）
#     1件として数える（利用の有無を見るため）
#   ・選定理由は、その祭り・その植物について記録された全部位の理由類型の和集合
# 【推定量】
#   raw_prev   素の割合          = 使用祭り数 / 30
#   w_prev     母集団推定割合    = Σ w_i·u_i / Σ w_i  （Σw_i = 156）
#   eq_prev    府県均等化割合    = 6府県の府県内使用率の単純平均
#                                  （母集団件数の推定を使わない感度分析用）
#   w_n        母集団での使用祭り数の推定値 = Σ w_i·u_i
#   CI         府県内で祭りを復元抽出する層化ブートストラップ（2000回）の
#              パーセンタイル区間。標本が府県3〜11件と小さいため区間は広い。
# ------------------------------------------------------------------------------

# 祭り単位の標本設計表（府県とウェイト）
fest_design <- tibble(festival = sheets) %>%
  mutate(pref = factor(unname(FESTIVAL_PREF[festival]), levels = PREF_ORDER)) %>%
  left_join(pref_weights(unname(FESTIVAL_PREF[sheets])), by = "pref")

# 祭り × 植物分類群（部位を統合、理由類型は和集合）
plant_festival <- resource_df %>%
  filter(!is.na(resource_taxon), resource_taxon != "非植物資材") %>%
  group_by(festival, resource_taxon) %>%
  summarise(
    n_parts      = n(),
    parts        = paste(unique(resource_raw), collapse = " / "),
    reason_types = {
      ty <- unique(unlist(strsplit(na.omit(reason_types), "\\|")))
      if (length(ty) == 0) NA_character_ else paste(sort(ty), collapse = "|")
    },
    .groups = "drop"
  ) %>%
  left_join(fest_design, by = "festival")

# --- 推定関数 ---------------------------------------------------------------
# u_mat: 祭り(行) × 植物(列) の 0/1 利用行列
usage_matrix <- function(df, unit_col) {
  tab <- table(factor(df$festival, levels = fest_design$festival),
               df[[unit_col]])
  (tab > 0) * 1
}

# 素・均等化・母集団推定の3種の割合をまとめて計算
estimate_prevalence <- function(u_mat, design = fest_design, n_boot = 2000, seed = 20260902) {
  w  <- design$w
  pr <- design$pref
  W  <- sum(w)
  raw_n  <- colSums(u_mat)
  w_n    <- as.numeric(t(u_mat) %*% w)
  # 府県均等化：府県内使用率の単純平均
  eq <- sapply(levels(pr), function(g) colMeans(u_mat[pr == g, , drop = FALSE]))
  eq_prev <- rowMeans(eq)

  set.seed(seed)
  idx_by_pref <- split(seq_len(nrow(u_mat)), pr)
  boot <- replicate(n_boot, {
    idx <- unlist(lapply(idx_by_pref, function(ii) sample(ii, length(ii), replace = TRUE)))
    as.numeric(t(u_mat[idx, , drop = FALSE]) %*% w[idx]) / W
  })
  tibble(
    taxon    = colnames(u_mat),
    raw_n    = as.numeric(raw_n),
    raw_prev = as.numeric(raw_n) / nrow(u_mat),
    eq_prev  = as.numeric(eq_prev),
    w_n      = w_n,
    w_prev   = w_n / W,
    lo       = apply(boot, 1, quantile, 0.025),
    hi       = apply(boot, 1, quantile, 0.975)
  )
}

# ==============================================================================
# 図01: 調査対象の基本プロファイル（協力者年齢 × 植物資源種数）
# ------------------------------------------------------------------------------
# 左：協力者の年齢（点＝個別協力者、横棒＝平均）
# 右：使用している植物資源の種類数（分類群ベース）
# 祭りの並び順は両図で共通。所属府県でグループ化し、府県内は平均年齢の高い順（上ほど高齢）。
# 右図は左図と同じ行に対応するため、縦軸のラベルを省略している。
# （旧 図04「植物資源種数」は本図の右パネルに統合した）
# ==============================================================================

# --- 植物資源の種類数（分類群ベース。図10・図12でも使用） ---
resource_diversity <- resource_df %>%
  group_by(festival) %>%
  summarise(n_resources = n_distinct(resource_taxon), .groups = "drop")

age_summary <- age_long %>%
  group_by(festival) %>%
  summarise(mean_age = mean(age), n_inf = n(), .groups = "drop")

# --- 全30祭りを含む並び順の土台（年齢・資源数が欠測の祭りも行を残す） ---
festival_profile <- tibble(festival = sheets) %>%
  mutate(pref = factor(unname(FESTIVAL_PREF[festival]), levels = PREF_ORDER)) %>%
  left_join(age_summary,        by = "festival") %>%
  left_join(resource_diversity, by = "festival")

if (any(is.na(festival_profile$pref))) {
  warning("府県未登録の祭り: ",
          paste(festival_profile$festival[is.na(festival_profile$pref)], collapse = ", "))
}

# 府県ごとにまとめ、府県内は平均年齢の昇順（coord_flip 後は上ほど高齢）
# coord_flip 後は最初の水準が下に来るため、昇順で並べると上ほど高齢になる
festival_order <- festival_profile %>%
  arrange(pref, mean_age) %>%
  pull(festival)

fct_fes <- function(x) factor(x, levels = festival_order)

festival_profile <- festival_profile %>% mutate(festival = fct_fes(festival))

age_long_f <- age_long %>%
  left_join(festival_profile %>% select(festival, pref) %>%
              mutate(festival = as.character(festival)), by = "festival") %>%
  mutate(festival = fct_fes(festival))

age_summary_f <- festival_profile %>% filter(!is.na(mean_age))

age_multi <- age_long_f %>%
  group_by(festival) %>%
  filter(n() > 1) %>%
  ungroup()

# 各祭りの協力者年齢を「53, 76, 77」の形にまとめた注記用ラベル
age_label_df <- age_long_f %>%
  filter(!is.na(age)) %>%
  group_by(festival, pref) %>%
  summarise(age_label = paste(sort(age), collapse = ", "), .groups = "drop")

# 府県ごとの行数に応じてパネル高さを可変にする共通設定
facet_pref <- facet_grid(pref ~ ., scales = "free_y", space = "free_y")

# --- 左：協力者年齢 ---
p01a_age <- ggplot(age_long_f, aes(x = festival, y = age)) +
  # 【重要】全30祭りを含む geom_blank を最初のレイヤーに置く。
  # これがないと離散軸は最初のレイヤー（複数協力者のみを含む age_multi）で
  # 訓練され、右パネルと行の並びがずれる。年齢が欠測の祭りの行もこれで残る。
  geom_blank(data = festival_profile, aes(x = festival, y = 60)) +
  geom_line(data = age_multi, aes(group = festival),
            color = "gray60", linewidth = 0.5) +
  geom_errorbar(data = age_summary_f,
                aes(x = festival, ymin = mean_age, ymax = mean_age),
                width = 0.6, color = "#444444", linewidth = 1.0,
                inherit.aes = FALSE) +
  geom_point(size = 3, alpha = 0.9, color = "#377EB8") +
  # 年齢の数値は点の脇ではなくパネル右端にまとめて表示（近い年齢の重なりを回避）
  geom_text(data = age_label_df, aes(x = festival, y = 99, label = age_label),
            hjust = 1, size = 2.9, color = "gray30", inherit.aes = FALSE) +
  coord_flip() +
  facet_pref +
  scale_y_continuous(limits = c(25, 100), breaks = seq(30, 90, 20)) +
  labs(title = "協力者年齢", x = NULL, y = "年齢（歳）") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        panel.grid.major.y = element_blank(),
        strip.text.y = element_blank(),
        strip.background.y = element_blank())

# --- 右：植物資源種数（縦軸ラベルなし） ---
p01b_div <- festival_profile %>%
  ggplot(aes(x = festival, y = n_resources)) +
  geom_col(fill = "#4DAF4A", alpha = 0.85, width = 0.7) +
  geom_text(aes(label = n_resources), hjust = -0.35, size = 2.9, color = "gray30") +
  coord_flip() +
  facet_pref +
  scale_y_continuous(limits = c(0, max(festival_profile$n_resources, na.rm = TRUE) + 1.5),
                     breaks = seq(0, 12, 2)) +
  labs(title = "植物資源種数", x = NULL, y = "植物資源の種類数（分類群）") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        panel.grid.major.y = element_blank(),
        axis.text.y = element_blank(),
        axis.ticks.y = element_blank())

p01_profile <- p01a_age + p01b_div +
  patchwork::plot_layout(widths = c(1, 0.62)) +
  patchwork::plot_annotation(
    title = "調査対象30火祭りの基本プロファイル",
    subtitle = paste0("左：協力者年齢（点＝個別協力者、縦線＝平均、右端の数字＝全協力者の年齢） ／ 右：植物資源の種類数",
                      "\n※祭りの並びは府県別、府県内は平均年齢の高い順"),
    caption = paste0("協力者 全", nrow(age_long), "名（複数名の祭りは点を縦線で連結）"),
    theme = theme(plot.title = element_text(face = "bold", family = "HiraginoSans-W3"),
                  plot.subtitle = element_text(family = "HiraginoSans-W3"),
                  plot.caption = element_text(family = "HiraginoSans-W3"))
  )

ggsave(file.path(OUTPUT_DIR, "01_profile_age_and_resources.png"), p01_profile,
       width = 11, height = 8.5, dpi = 150)

# プロファイル表（府県・平均年齢・資源種数）
write.csv(
  festival_profile %>%
    arrange(pref, mean_age) %>%
    select(pref, festival, mean_age, n_informants = n_inf, n_resources),
  file.path(OUTPUT_DIR, "festival_profile.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

# ==============================================================================
# 図02: 関係者数・観光客数トレンド（横ばい = 黄色）
# ==============================================================================

trend_order  <- c("大幅増加", "増加", "横ばい", "減少", "大幅減少", "なし/0", "不明")
trend_colors <- c(
  "大幅増加" = "#1B7837",
  "増加"     = "#7FBF7B",
  "横ばい"   = "#FFD700",   # 黄色
  "減少"     = "#F4A582",
  "大幅減少" = "#D6604D",
  "なし/0"   = "#BDBDBD",
  "不明"     = "#E0E0E0"
)

trend_long <- bind_rows(
  survey_df %>% count(trend = participants_trend) %>% mutate(type = "祭り関係者数"),
  survey_df %>% count(trend = tourists_trend)     %>% mutate(type = "観光客数")
) %>%
  mutate(trend = factor(trend, levels = trend_order))

p02_trend <- trend_long %>%
  ggplot(aes(x = type, y = n, fill = trend)) +
  geom_col(position = "stack", width = 0.5) +
  geom_text(aes(label = n), position = position_stack(vjust = 0.5),
            size = 3.5, color = "gray20") +
  scale_fill_manual(values = trend_colors, name = "変化方向",
                    drop = FALSE) +
  labs(title = "関係者数・観光客数の変化傾向",
       subtitle = paste0("全", nrow(survey_df), "祭り"),
       x = NULL, y = "祭り件数") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "right")

ggsave(file.path(OUTPUT_DIR, "02_trend_participants_tourists.png"), p02_trend,
       width = 7, height = 5, dpi = 150)

# ==============================================================================
# 図03: 植物の利用頻度 — 抽出の府県偏りを補正した推定
# ------------------------------------------------------------------------------
# 03a 素の利用祭り数（部位重複を除いた祭り数）と母集団推定割合の比較
# 03b 府県別の利用率ヒートマップ（順位が動く理由を示す）
# ==============================================================================

u_taxon  <- usage_matrix(plant_festival, "resource_taxon")
prev_tax <- estimate_prevalence(u_taxon)

cat("\n=== 府県別の抽出率とウェイト ===\n")
print(as.data.frame(pref_weights(unname(FESTIVAL_PREF[sheets]))))

cat("\n=== 植物の利用頻度（素 vs 補正）===\n")
print(as.data.frame(
  prev_tax %>%
    arrange(desc(w_prev)) %>%
    transmute(taxon, raw_n,
              素の割合 = round(raw_prev, 3),
              府県均等化 = round(eq_prev, 3),
              母集団推定割合 = round(w_prev, 3),
              CI = paste0("[", round(lo, 2), ", ", round(hi, 2), "]"),
              推定使用祭り数 = round(w_n, 1))
))

# --- 03a: 素 vs 母集団推定 --------------------------------------------------
prev_plot <- prev_tax %>%
  filter(raw_n >= 2) %>%                     # 1祭りのみの植物は推定が不安定なため除外
  mutate(taxon = fct_reorder(taxon, w_prev))

p03a <- ggplot(prev_plot, aes(y = taxon)) +
  geom_errorbar(aes(xmin = lo, xmax = hi), orientation = "y", width = 0,
                color = "gray70", linewidth = 0.8) +
  geom_segment(aes(x = raw_prev, xend = w_prev, y = taxon, yend = taxon),
               color = "gray45", linewidth = 0.4,
               arrow = arrow(length = unit(0.10, "cm"), type = "closed")) +
  geom_point(aes(x = raw_prev, color = "素の割合（30祭りのうち）"), size = 2.6) +
  geom_point(aes(x = w_prev,   color = "母集団推定割合（府県ウェイト）"), size = 3.2) +
  scale_color_manual(values = c("素の割合（30祭りのうち）" = "#999999",
                                "母集団推定割合（府県ウェイト）" = "#D62728"),
                     name = NULL) +
  scale_x_continuous(labels = scales::percent, limits = c(0, 1)) +
  labs(
    title = "植物ごとの利用の広がり — 抽出の府県偏りを補正",
    subtitle = paste0("解析単位＝祭り×植物分類群（同一祭り内の部位重複は1件に集約） ／ ",
                      "2祭り以上で使用された植物のみ\n",
                      "灰＝素の割合、赤＝母集団156件に事後層化した推定割合、横棒＝層化ブートストラップ95%区間"),
    x = "その植物を使用する火祭りの割合", y = NULL,
    caption = "矢印は補正による移動方向。区間は府県あたり標本が小さいため広い。"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "bottom",
        panel.grid.major.y = element_blank())

ggsave(file.path(OUTPUT_DIR, "03a_plant_prevalence_weighted.png"), p03a,
       width = 9.5, height = max(5, nrow(prev_plot) * 0.42), dpi = 150)

# --- 03b: 府県別の利用率ヒートマップ ----------------------------------------
pref_prev <- plant_festival %>%
  count(pref, resource_taxon) %>%
  right_join(expand_grid(pref = factor(PREF_ORDER, levels = PREF_ORDER),
                         resource_taxon = prev_plot$taxon),
             by = c("pref", "resource_taxon")) %>%
  mutate(n = ifelse(is.na(n), 0, n)) %>%
  left_join(pref_weights(unname(FESTIVAL_PREF[sheets])) %>% select(pref, n_sample),
            by = "pref") %>%
  mutate(prev = n / n_sample,
         resource_taxon = factor(resource_taxon, levels = levels(prev_plot$taxon)))

p03b <- ggplot(pref_prev, aes(x = pref, y = resource_taxon, fill = prev)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(n > 0, paste0(n, "/", n_sample), "")),
            size = 2.8, color = "gray20", family = "HiraginoSans-W3") +
  scale_fill_gradient(low = "#F7FBFF", high = "#08519C",
                      labels = scales::percent, name = "府県内の利用率") +
  labs(title = "府県別の植物利用率",
       subtitle = paste0("セル内は「使用した祭り数／その府県の調査祭り数」。\n",
                         "府県ごとの標本数の差が図03aの補正量を決める。"),
       x = NULL, y = NULL) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        panel.grid = element_blank(),
        legend.position = "bottom")

ggsave(file.path(OUTPUT_DIR, "03b_plant_prevalence_by_pref.png"), p03b,
       width = 8.5, height = max(5, nrow(prev_plot) * 0.36), dpi = 150)

write.csv(
  prev_tax %>% arrange(desc(w_prev)),
  file.path(OUTPUT_DIR, "plant_prevalence_weighted.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

# ==============================================================================
# 図04: 植物資源種数（祭り別）→ 図01 の右パネルに統合（2026-09-02）
# resource_diversity の計算は図01のブロックに移動した。
# ==============================================================================

# ==============================================================================
# 図05: 組織の有無
# ==============================================================================

org_summary <- survey_df %>%
  mutate(
    `保存会あり`         = !is.na(preservation) &
                           !str_detect(str_replace_all(preservation, "\\s+",""), "^なし$"),
    `氏子組織あり`       = !is.na(ujiko_org) &
                           !str_detect(str_replace_all(ujiko_org, "\\s+",""), "^なし$"),
    `行政・自治組織あり` = !is.na(admin_org) &
                           !str_detect(str_replace_all(admin_org, "\\s+",""), "^なし$"),
    `企業・財団あり`     = !is.na(company) &
                           !str_detect(str_replace_all(company, "\\s+",""), "^なし$"),
    `学校あり`           = !is.na(school) &
                           !str_detect(str_replace_all(school, "\\s+",""), "^なし$")
  ) %>%
  summarise(across(ends_with("あり"), sum)) %>%
  pivot_longer(everything(), names_to = "組織種別", values_to = "件数")

p05_org <- org_summary %>%
  mutate(組織種別 = fct_reorder(組織種別, 件数)) %>%
  ggplot(aes(x = 組織種別, y = 件数)) +
  geom_col(fill = "#984EA3", alpha = 0.85) +
  geom_text(aes(label = paste0(件数, "/", nrow(survey_df))),
            hjust = -0.1, size = 3.5) +
  coord_flip() +
  scale_y_continuous(limits = c(0, nrow(survey_df) + 1)) +
  labs(title = "火祭りに関わる組織の有無",
       x = NULL, y = "件数") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"))

ggsave(file.path(OUTPUT_DIR, "05_organization_presence.png"), p05_org,
       width = 8, height = 4, dpi = 150)

# ==============================================================================
# 図06: 祭り目的キーワード
# ==============================================================================

purpose_kw <- c(
  "五穀豊穣|豊作|農業" = "豊作祈願",
  "無病息災"           = "無病息災",
  "家内安全"           = "家内安全",
  "地域安全"           = "地域安全",
  "先祖供養"           = "先祖供養",
  "伝統.*継承|文化.*継承" = "伝統継承",
  "火伏|火除"          = "火伏せ/火除け",
  "神恩感謝"           = "神恩感謝",
  "商売繁盛"           = "商売繁盛",
  "疫病退散"           = "疫病退散",
  "環境保護|水質"      = "環境保護",
  "交通安全"           = "交通安全"
)

purpose_count <- lapply(names(purpose_kw), function(pat) {
  data.frame(keyword = purpose_kw[[pat]],
             n = sum(str_detect(survey_df$purpose, regex(pat)), na.rm = TRUE))
}) %>% bind_rows() %>% filter(n > 0) %>% arrange(desc(n))

p06_purpose <- purpose_count %>%
  mutate(keyword = fct_reorder(keyword, n)) %>%
  ggplot(aes(x = keyword, y = n)) +
  geom_col(fill = "#E6550D", alpha = 0.85) +
  geom_text(aes(label = n), hjust = -0.2, size = 3.5) +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(purpose_count$n) + 0.5)) +
  labs(title = "祭り目的のキーワード出現頻度",
       x = NULL, y = "件数") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"))

ggsave(file.path(OUTPUT_DIR, "06_festival_purpose.png"), p06_purpose,
       width = 8, height = 5, dpi = 150)

# ==============================================================================
# 図07: 信仰キーワード
# ==============================================================================

belief_kw <- c(
  "氏神"              = "氏神信仰",
  "愛宕"              = "愛宕信仰",
  "祖霊|お盆"         = "祖霊/お盆信仰",
  "八幡"              = "八幡信仰",
  "春祭"              = "春祭信仰",
  "琵琶湖"            = "琵琶湖信仰",
  "伝統.*継承|文化遺産" = "伝統継承（現代）",
  "歳神"              = "歳神信仰",
  "不動明王|愛宕"     = "仏教系信仰",
  "稲荷"              = "稲荷信仰"
)

belief_count <- lapply(names(belief_kw), function(pat) {
  data.frame(keyword = belief_kw[[pat]],
             n = sum(str_detect(survey_df$belief, regex(pat)), na.rm = TRUE))
}) %>% bind_rows() %>% distinct(keyword, .keep_all = TRUE) %>%
  filter(n > 0) %>% arrange(desc(n))

p07_belief <- belief_count %>%
  mutate(keyword = fct_reorder(keyword, n)) %>%
  ggplot(aes(x = keyword, y = n)) +
  geom_col(fill = "#3182BD", alpha = 0.85) +
  geom_text(aes(label = n), hjust = -0.2, size = 3.5) +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(belief_count$n) + 0.5)) +
  labs(title = "信仰の種類と出現頻度",
       x = NULL, y = "件数") +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"))

ggsave(file.path(OUTPUT_DIR, "07_belief_types.png"), p07_belief,
       width = 8, height = 5, dpi = 150)

# ==============================================================================
# 図08: 近年の課題（自由記述 → キーワード分類）
# ------------------------------------------------------------------------------
# 【処理方針の注記】
# 原問卷の回答形式は自由記述であり、祭りによって粒度が大きく異なる。
#   - 短答型：「若者が少ない」「ない」など1〜2語
#   - 長文型：王の浜・小田神社・大嶋奥津嶋神社は3〜5文の段落で記述
# 処理方法：正規表現によるキーワードマッチング（主題コーディング）
#   各キーワードパターンに対して str_detect() でマッチするかをチェック。
#   1件の回答が複数のカテゴリにヒットすることがある（例：小田神社は
#   「若者不足」「費用」「資源調達」を同時に言及）。
#   → グラフは「件数」ではなく「延べ言及件数」になる点に注意。
#   長文型の回答は完全にはカバーできないため、必要に応じて
#   human coding との照合推奨。
# ==============================================================================

challenge_kw <- c(
  "若者|若年|後継|担い手|参加者.*減|人手"       = "若者不足・担い手問題",
  "少子化|高齢化|高齢|老年"                     = "少子高齢化",
  "資源|植物|調達|入手|刈|購入|材料"             = "植物資源の調達難",
  "費用|資金|コスト|経費|お金|財政"             = "費用・資金問題",
  "コロナ|COVID|疫病|感染"                       = "コロナ禍の影響",
  "技術|技法|作り方|製法|伝承方法"               = "技術継承",
  "外来種|環境|生態|消滅|減少.*植"               = "植生環境の変化",
  "過疎|人口減|転出|農家.*減"                    = "過疎化・人口減少"
)

challenge_count <- lapply(names(challenge_kw), function(pat) {
  data.frame(
    keyword = challenge_kw[[pat]],
    n = sum(str_detect(survey_df$challenges, regex(pat, ignore_case = TRUE)),
            na.rm = TRUE)
  )
}) %>% bind_rows() %>% filter(n > 0) %>% arrange(desc(n))

cat("\n=== 近年の課題（キーワード分類 — 延べ言及件数） ===\n")
print(challenge_count)

p08_challenge <- challenge_count %>%
  mutate(keyword = fct_reorder(keyword, n)) %>%
  ggplot(aes(x = keyword, y = n)) +
  geom_col(fill = "#D62728", alpha = 0.85) +
  geom_text(aes(label = n), hjust = -0.2, size = 3.5) +
  coord_flip() +
  scale_y_continuous(limits = c(0, max(challenge_count$n) + 0.5)) +
  labs(
    title = "近年の課題（キーワード頻度）",
    subtitle = "自由記述をキーワードで分類（1件が複数カテゴリに該当する場合あり）",
    x = NULL, y = "言及件数（延べ）"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        plot.subtitle = element_text(size = 8, color = "gray40"))

ggsave(file.path(OUTPUT_DIR, "08_recent_challenges.png"), p08_challenge,
       width = 9, height = 5, dpi = 150)

# ==============================================================================
# 図09a: 植物資源 × 祭り — 代替可能性マトリクス
# 図09b: 植物資源 × 祭り — 調達嵌入度マトリクス
# ==============================================================================

# 図03の resource_count は廃止。2祭り以上で使われた資源名（記録どおりの粒度）
resource_count <- resource_df %>%
  distinct(festival, resource_norm) %>%
  count(resource_norm, sort = TRUE)
top_resources <- resource_count %>% filter(n >= 2) %>% pull(resource_norm)

matrix_theme <- theme_bw(base_family = "HiraginoSans-W3") +
  theme(
    plot.title  = element_text(face = "bold"),
    axis.text.x = element_text(angle = 45, hjust = 1),
    axis.text.y = element_text(size = 8),
    panel.grid  = element_blank(),
    legend.position = "right"
  )

# --- 09a: 代替可能性（1=代替可・緑, 3=代替不可・赤, NA=灰=不使用） ---
matrix_subst <- resource_df %>%
  filter(resource_norm %in% top_resources, !is.na(subst_score)) %>%
  group_by(festival, resource_norm) %>%
  summarise(subst_score = mean(subst_score, na.rm = TRUE), .groups = "drop") %>%
  complete(festival = sheets, resource_norm = top_resources)

p09a_subst <- matrix_subst %>%
  ggplot(aes(x = resource_norm, y = festival, fill = subst_score)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(!is.na(subst_score),
                               c("可","困","不")[round(subst_score)], "")),
            size = 2.5, color = "white", fontface = "bold") +
  scale_fill_gradient(low = "#74C476", high = "#D62728", na.value = "#EEEEEE",
                      limits = c(1, 3),
                      breaks = 1:3,
                      labels = c("1 代替可", "2 困難", "3 不可"),
                      name = "代替可能性") +
  labs(title = "植物資源の使用状況 × 代替可能性",
       subtitle = "灰色 = 当該祭りで使用なし",
       x = NULL, y = NULL) +
  matrix_theme

ggsave(file.path(OUTPUT_DIR, "09a_resource_matrix_substitutability.png"),
       p09a_subst, width = 11, height = 7, dpi = 150)

# --- 09b: 調達嵌入度（1=購入・赤, 3=自採・緑, NA=灰=不使用） ---
matrix_embed <- resource_df %>%
  filter(resource_norm %in% top_resources, !is.na(embed_score)) %>%
  group_by(festival, resource_norm) %>%
  summarise(embed_score = mean(embed_score, na.rm = TRUE), .groups = "drop") %>%
  complete(festival = sheets, resource_norm = top_resources)

p09b_embed <- matrix_embed %>%
  ggplot(aes(x = resource_norm, y = festival, fill = embed_score)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(!is.na(embed_score),
                               c("購","農","採")[round(embed_score)], "")),
            size = 2.5, color = "white", fontface = "bold") +
  scale_fill_gradient(low = "#D62728", high = "#1A6A1A", na.value = "#EEEEEE",
                      limits = c(1, 3),
                      breaks = 1:3,
                      labels = c("1 市場購入", "2 地域農家", "3 共同体自採"),
                      name = "調達嵌入度") +
  labs(title = "植物資源の使用状況 × 調達嵌入度",
       subtitle = "灰色 = 当該祭りで使用なし",
       x = NULL, y = NULL) +
  matrix_theme

ggsave(file.path(OUTPUT_DIR, "09b_resource_matrix_embeddedness.png"),
       p09b_embed, width = 11, height = 7, dpi = 150)

# ==============================================================================
# 図10: 関係者数 vs 観光客数 — 散布図
# ------------------------------------------------------------------------------
# 【分析の注記】
# 原データは文字列混じりのため、以下の方針で数値を手動設定。
#   - 関係者数：複数記載があれば最大値を採用（保存会+氏子+ボランティア等）
#   - 観光客数：ライブ配信等の間接視聴は除外し、現地参集人数を採用
#   - 「不明」はNA扱い（グラフ外）
# ==============================================================================

scale_df <- tribble(
  ~festival,                    ~participants, ~tourists,
  "巽神社松明",                  36,           0,
  "鞍馬の火祭",                  136,          5000,
  "三栖の火祭",                  150,          3500,
  "松明を次世代に送る会",        1000,         0,
  "雄琴学区ヨシ松明一斉点火",    500,          NA,    # 不明
  "太郎坊宮の火祭り",            100,          0,
  "信楽の火祭り",                180,          3000,
  "勝部の火祭り",                150,          2500,  # 現地のみ
  "近江八幡左義長祭り",          600,          80000,
  "大文字送り火",                700,          350000,
  "八幡祭り",                    550,          16500,
  "王の浜若宮神社",              30,           30,
  "小田神社",                    60,           400,
  "大嶋奥津嶋神社",              50,           20,
  # ---- 2026-09-02 追加：30シート版の新規16祭り ------------------------------
  # 上と同じ方針（関係者数=複数記載なら最大値、観光客数=現地参集人数）で
  # 自由記述から読み取った値。★ は原文が幅・感覚値・未計数のため要確認。
  "往馬大社",                    200,          5000,
  "熊野速玉大社",                2000,         NA,     # ★ 運営100人＋上り子・本番参加者1,500〜2,000人／観光客数の明示なし
  "稲むらの火祭り",              200,          600,    # ★ 「参加者600人程」＝見物客とほぼ同義との説明
  "熊野那智",                    120,          3000,
  "嵯峨のお松明式",              40,           1000,   # ★ 「1,000人を大きく上回る」正確な計数なし → 下限値
  "広河原松上げ",                80,           1000,
  "がんがら火祭り",              53,           10000,  # ★ 関係者数はLINE登録53名のみ（当日協力者は未計数）
  "まんどろ火祭り",              16,           4000,   # ★ 関係者数は実行委員会15〜16名のみ
  "麦わら松明",                  10,           100,    # ★ 中核10人前後／観光客は未計数「100人以上」
  "東光寺鬼会",                  21,           200,
  "吉祥草寺茅原大とんど",        170,          3500,
  "稲引き樽引き神事",            20,           60,
  "花背松上げ",                  30,           300,    # ★ 「二、三百人」「五百人ぐらい」複数の感覚値
  "雲ケ畑松上げ",                10,           30,
  "湯村火祭り",                  53,           300,
  "ほうらんや火祭り",            10,           100     # ★ 代表層約10名のみ／2026年は神事のみで大幅減
) %>%
  left_join(resource_diversity, by = "festival")

# 祭り分類 2×2：
#   関係者数 ≥ 200 → 大規模、< 200 → 小規模
#   観光客数 ≥ 1000 → 観光集客型、< 1000（0含む）→ 地域内向型
#   雄琴学区（観光客数不明）は関係者数のみで判定 → 大規模・地域型とみなす
scale_df <- scale_df %>%
  mutate(
    participants_class = ifelse(participants >= 200, "大規模", "小規模"),
    tourists_class     = case_when(
      is.na(tourists)    ~ "地域型",   # 不明は地域型扱い
      tourists >= 1000   ~ "観光型",
      TRUE               ~ "地域型"
    ),
    festival_type = paste0(participants_class, "・", tourists_class)
  )

type_colors <- c(
  "大規模・観光型" = "#D62728",   # 赤：大文字・左義長・八幡
  "大規模・地域型" = "#FF7F0E",   # 橙：松明を次世代・雄琴
  "小規模・観光型" = "#9467BD",   # 紫：鞍馬・三栖・信楽・勝部
  "小規模・地域型" = "#1F77B4"    # 青：巽神社・太郎坊・王の浜・小田・大嶋
)

p10_scatter <- scale_df %>%
  filter(!is.na(tourists)) %>%
  ggplot(aes(x = participants, y = tourists + 1,
             color = festival_type, size = n_resources)) +
  # 2×2 象限の区切り線
  geom_vline(xintercept = 200, linetype = "dashed", color = "gray60", linewidth = 0.5) +
  geom_hline(yintercept = 1001, linetype = "dashed", color = "gray60", linewidth = 0.5) +
  # 象限ラベル
  annotate("text", x = 80,  y = 500000, label = "小規模・観光型", size = 2.8,
           color = "#9467BD", alpha = 0.7, fontface = "italic",
           family = "HiraginoSans-W3") +
  annotate("text", x = 800, y = 500000, label = "大規模・観光型", size = 2.8,
           color = "#D62728", alpha = 0.7, fontface = "italic",
           family = "HiraginoSans-W3") +
  annotate("text", x = 80,  y = 2,      label = "小規模・地域型", size = 2.8,
           color = "#1F77B4", alpha = 0.7, fontface = "italic",
           family = "HiraginoSans-W3") +
  annotate("text", x = 800, y = 2,      label = "大規模・地域型", size = 2.8,
           color = "#FF7F0E", alpha = 0.7, fontface = "italic",
           family = "HiraginoSans-W3") +
  geom_point(alpha = 0.85) +
  geom_text_repel(aes(label = festival), size = 2.8, max.overlaps = 20,
                  family = "HiraginoSans-W3") +
  scale_x_continuous(transform = "log10", labels = scales::comma,
                     breaks = c(30, 100, 200, 500, 1000)) +
  scale_y_continuous(transform = "log10",
                     breaks = c(1, 10, 100, 1000, 10000, 100000, 1000000),
                     labels = c("0", "10", "100", "1千", "1万", "10万", "100万")) +
  scale_color_manual(values = type_colors, name = "祭りタイプ（2×2）") +
  scale_size_continuous(range = c(3, 9), name = "植物資源種数") +
  labs(
    title = "関係者数 vs 観光客数（2×2 類型）",
    subtitle = "縦破線: 関係者数200人、横破線: 観光客1000人 | 点の大きさ = 植物資源種数",
    x = "祭り関係者数（人）[対数]",
    y = "観光客数（人）[対数]",
    caption = "雄琴学区ヨシ松明一斉点火・熊野速玉大社は観光客数不明のため除外"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "right")

ggsave(file.path(OUTPUT_DIR, "10_participants_vs_tourists.png"), p10_scatter,
       width = 10, height = 7, dpi = 150)

# 相関係数（対数変換後）
cor_data <- scale_df %>%
  filter(!is.na(tourists) & tourists > 0 & participants > 0)
if (nrow(cor_data) >= 4) {
  cor_result <- cor.test(log10(cor_data$participants), log10(cor_data$tourists),
                         method = "pearson")
  cat("\n=== 関係者数 vs 観光客数 相関（対数変換, n =", nrow(cor_data), ") ===\n")
  cat("Pearson r =", round(cor_result$estimate, 3),
      ", p =", round(cor_result$p.value, 3), "\n")
}

# 図11は削除

# ==============================================================================
# 図12: 生息地多様性 vs 植物資源種数
# ------------------------------------------------------------------------------
# 生息地多様性 = 各祭りで使用される植物の調達先景観タイプの数
# （水田・湿地・森林・畑・荒地・庭などのユニーク数）
# ==============================================================================

habitat_diversity <- resource_df %>%
  filter(!is.na(landscape_norm) & landscape_norm != "なし") %>%
  group_by(festival) %>%
  summarise(
    n_habitats  = n_distinct(landscape_norm),
    habitats    = paste(sort(unique(landscape_norm)), collapse = "・"),
    .groups = "drop"
  )

cat("\n=== 生息地多様性（祭り別） ===\n")
print(habitat_diversity %>% arrange(desc(n_habitats)))

div_join <- resource_diversity %>%
  left_join(habitat_diversity, by = "festival") %>%
  left_join(scale_df %>% select(festival, festival_type), by = "festival")

p12_habitat <- div_join %>%
  filter(!is.na(n_habitats)) %>%
  ggplot(aes(x = n_habitats, y = n_resources,
             color = festival_type)) +
  geom_jitter(size = 4, alpha = 0.85, width = 0.05, height = 0.05) +
  geom_smooth(method = "lm", se = TRUE, color = "gray50",
              linetype = "dashed", linewidth = 0.8, inherit.aes = FALSE,
              aes(x = n_habitats, y = n_resources), data = div_join) +
  geom_text_repel(aes(label = festival), size = 2.8, max.overlaps = 20,
                  family = "HiraginoSans-W3") +
  scale_x_continuous(breaks = 1:6) +
  scale_color_manual(values = type_colors, name = "祭りタイプ") +
  labs(
    title = "生息地多様性と植物資源種数の関係",
    subtitle = "生息地多様性 = 調達先景観タイプの数（水田・湿地・森林等）",
    x = "生息地タイプ数（景観の多様性）",
    y = "植物資源種数"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"))

ggsave(file.path(OUTPUT_DIR, "12_habitat_diversity_vs_resources.png"),
       p12_habitat, width = 9, height = 6, dpi = 150)

cor12 <- cor.test(div_join$n_habitats, div_join$n_resources,
                  method = "pearson", use = "complete.obs")
cat("\n=== 生息地多様性 vs 植物資源種数 相関 ===\n")
cat("r =", round(cor12$estimate, 3), ", p =", round(cor12$p.value, 3), "\n")

# ==============================================================================
# 図13: 植物資源 × 信仰 クロス分析
# ==============================================================================

# 信仰カテゴリを各祭りに付与
belief_cats <- c(
  "氏神"              = "氏神信仰",
  "愛宕"              = "愛宕信仰",
  "祖霊|お盆"         = "祖霊/お盆信仰",
  "八幡"              = "八幡信仰",
  "琵琶湖"            = "琵琶湖信仰",
  "伝統.*継承|文化遺産" = "伝統継承（現代）"
)

festival_belief <- survey_df %>%
  select(festival, belief) %>%
  rowwise() %>%
  mutate(
    belief_cats = list(
      names(belief_cats)[
        sapply(names(belief_cats),
               function(pat) str_detect(belief %||% "", regex(pat, ignore_case = TRUE)))
      ]
    ),
    belief_label = paste(belief_cats[belief_cats != ""], collapse = " / ")
  ) %>%
  ungroup() %>%
  unnest(belief_cats) %>%
  filter(belief_cats != "") %>%
  mutate(belief_label_mapped = belief_cats[belief_cats])

# 資源 × 信仰の共起マトリクス
resource_belief <- resource_df %>%
  filter(resource_norm %in% top_resources) %>%
  distinct(festival, resource_norm) %>%
  left_join(festival_belief %>% select(festival, belief_cats), by = "festival") %>%
  filter(!is.na(belief_cats)) %>%
  group_by(resource_norm, belief_cats) %>%
  summarise(n = n(), .groups = "drop")

p13_belief <- resource_belief %>%
  ggplot(aes(x = belief_cats, y = resource_norm, fill = n)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(n > 0, n, "")), size = 3) +
  scale_fill_gradient(low = "#EFF3FF", high = "#08519C",
                      name = "共起件数") +
  labs(
    title = "植物資源 × 信仰 共起マトリクス",
    subtitle = "セル値 = その信仰をもつ祭りで当該植物資源が使われた件数",
    x = "信仰カテゴリ", y = "植物資源"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title  = element_text(face = "bold"),
        axis.text.x = element_text(angle = 30, hjust = 1),
        panel.grid  = element_blank())

ggsave(file.path(OUTPUT_DIR, "13_resource_x_belief.png"), p13_belief,
       width = 10, height = 7, dpi = 150)

# ==============================================================================
# 図14: 植物資源 × 祭り目的 クロス分析
# ==============================================================================

purpose_cats <- c(
  "五穀豊穣|豊作"      = "豊作祈願",
  "無病息災"           = "無病息災",
  "家内安全"           = "家内安全",
  "先祖供養"           = "先祖供養",
  "伝統.*継承|文化.*継承" = "伝統継承",
  "火伏|火除"          = "火伏せ/火除け",
  "環境保護|水質"      = "環境保護",
  "地域安全"           = "地域安全"
)

festival_purpose <- survey_df %>%
  select(festival, purpose) %>%
  rowwise() %>%
  mutate(
    purpose_cats = list(
      names(purpose_cats)[
        sapply(names(purpose_cats),
               function(pat) str_detect(purpose %||% "", regex(pat, ignore_case = TRUE)))
      ]
    )
  ) %>%
  ungroup() %>%
  unnest(purpose_cats) %>%
  filter(purpose_cats != "") %>%
  mutate(purpose_label = purpose_cats[purpose_cats])

resource_purpose <- resource_df %>%
  filter(resource_norm %in% top_resources) %>%
  distinct(festival, resource_norm) %>%
  left_join(festival_purpose %>% select(festival, purpose_cats), by = "festival") %>%
  filter(!is.na(purpose_cats)) %>%
  group_by(resource_norm, purpose_cats) %>%
  summarise(n = n(), .groups = "drop")

p14_purpose <- resource_purpose %>%
  ggplot(aes(x = purpose_cats, y = resource_norm, fill = n)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(n > 0, n, "")), size = 3) +
  scale_fill_gradient(low = "#FFF5EB", high = "#D94801",
                      name = "共起件数") +
  labs(
    title = "植物資源 × 祭り目的 共起マトリクス",
    subtitle = "セル値 = その目的をもつ祭りで当該植物資源が使われた件数",
    x = "祭り目的カテゴリ", y = "植物資源"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title  = element_text(face = "bold"),
        axis.text.x = element_text(angle = 30, hjust = 1),
        panel.grid  = element_blank())

ggsave(file.path(OUTPUT_DIR, "14_resource_x_purpose.png"), p14_purpose,
       width = 10, height = 7, dpi = 150)

# ==============================================================================
# CSV出力
# ==============================================================================

# ==============================================================================
# 図15: 文化的関鍵種（Cultural Keystone Species）
#   x軸 = 出現頻度（祭り数）、y軸 = 代替可能性スコア平均
#   右上ほど「多くの祭りで使われ、かつ代替不可」= 文化的関鍵種
# ==============================================================================

keystone_df <- resource_df %>%
  filter(!is.na(subst_score)) %>%
  group_by(resource_norm) %>%
  summarise(
    n_festivals  = n_distinct(festival),
    mean_subst   = mean(subst_score, na.rm = TRUE),
    mean_embed   = mean(embed_score, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    keystone_score = n_festivals * mean_subst,
    label_flag = mean_subst >= 2.5 | n_festivals >= 4
  )

cat("\n=== 文化的関鍵種候補（代替不可能性 × 出現頻度） ===\n")
print(as.data.frame(keystone_df %>% arrange(desc(keystone_score))))

p15_keystone <- keystone_df %>%
  ggplot(aes(x = n_festivals, y = mean_subst, size = mean_embed)) +
  # 背景ゾーン（右上 = 関鍵種ゾーン）
  annotate("rect", xmin = 3.5, xmax = Inf, ymin = 2.4, ymax = Inf,
           fill = "#FFF3CD", alpha = 0.6) +
  annotate("text", x = 4, y = 3.05, label = "文化的関鍵種ゾーン",
           hjust = 0, size = 3, color = "#856404", fontface = "italic",
           family = "HiraginoSans-W3") +
  geom_point(aes(color = mean_subst), alpha = 0.85) +
  geom_text_repel(
    aes(label = resource_norm), size = 3,
    family = "HiraginoSans-W3", max.overlaps = 30
  ) +
  scale_color_gradient(low = "#74C476", high = "#D62728",
                       name = "代替不可能性\n（平均スコア）") +
  scale_size_continuous(range = c(2, 8), name = "調達嵌入度\n（平均スコア）") +
  scale_x_continuous(breaks = 1:max(keystone_df$n_festivals)) +
  scale_y_continuous(limits = c(0.8, 3.2), breaks = 1:3,
                     labels = c("1\n代替可", "2\n代替困難", "3\n代替不可")) +
  labs(
    title = "文化的関鍵植物種（Cultural Keystone Species）",
    subtitle = "右上ほど「広く使われ、かつ代替できない」文化的に不可欠な種",
    x = "利用祭り数",
    y = "代替可能性スコア（平均）"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "right")

ggsave(file.path(OUTPUT_DIR, "15_cultural_keystone_species.png"), p15_keystone,
       width = 10, height = 7, dpi = 150)

# ==============================================================================
# 図16: 文化-生態嵌入度（調達方法の地域性）
#   各祭りの嵌入度分布：3=共同体自採、2=地域農家から、1=市場購入
# ==============================================================================

embed_festival <- resource_df %>%
  filter(!is.na(embed_score)) %>%
  group_by(festival) %>%
  summarise(
    mean_embed   = mean(embed_score, na.rm = TRUE),
    n_high       = sum(embed_score == 3, na.rm = TRUE),
    n_mid        = sum(embed_score == 2, na.rm = TRUE),
    n_low        = sum(embed_score == 1, na.rm = TRUE),
    n_total      = n(),
    .groups = "drop"
  ) %>%
  arrange(desc(mean_embed))

cat("\n=== 文化-生態嵌入度（祭り別） ===\n")
print(as.data.frame(embed_festival))

embed_long <- resource_df %>%
  filter(!is.na(embed_score)) %>%
  mutate(
    embed_label = factor(embed_score,
                         levels = c(3, 2, 1),
                         labels = c("3: 共同体自採/自栽培",
                                    "2: 地域農家・知人から",
                                    "1: 市場・業者購入"))
  ) %>%
  group_by(festival, embed_label) %>%
  summarise(n = n(), .groups = "drop") %>%
  group_by(festival) %>%
  mutate(pct = n / sum(n)) %>%
  ungroup()

festival_embed_order <- embed_festival %>% pull(festival)

p16_embed <- embed_long %>%
  mutate(festival = factor(festival, levels = festival_embed_order)) %>%
  ggplot(aes(x = festival, y = pct, fill = embed_label)) +
  geom_col(position = "stack", width = 0.7) +
  geom_text(
    data = embed_festival %>%
      mutate(festival = factor(festival, levels = festival_embed_order)),
    aes(x = festival, y = 1.05, label = sprintf("%.1f", mean_embed)),
    inherit.aes = FALSE, size = 3, color = "gray30"
  ) +
  coord_flip() +
  scale_fill_manual(
    values = c("3: 共同体自採/自栽培"  = "#1A6A1A",
               "2: 地域農家・知人から" = "#74C476",
               "1: 市場・業者購入"     = "#D62728"),
    name = "調達方法"
  ) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.12)) +
  labs(
    title = "文化-生態嵌入度：植物資源の調達方法（祭り別）",
    subtitle = "右の数値 = 嵌入度スコア平均（3=最高、1=最低）",
    x = NULL, y = "植物資源の割合"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "bottom")

ggsave(file.path(OUTPUT_DIR, "16_procurement_embeddedness.png"), p16_embed,
       width = 10, height = 7, dpi = 150)

# ==============================================================================
# 図17: 植物の選定理由の類型（旧TEK分析を全面改訂）
# ------------------------------------------------------------------------------
# 17a 理由類型の全体分布（素 vs 府県ウェイト補正）
# 17b 主要植物ごとの理由構成
# 解析単位は 祭り×植物分類群。1単位が複数の理由類型を持つため、
# 割合の合計は100%を超える（多重ラベル）。
# ==============================================================================

reason_long <- plant_festival %>%
  filter(!is.na(reason_types)) %>%
  separate_rows(reason_types, sep = "\\|") %>%
  rename(rtype = reason_types) %>%
  mutate(rlabel = factor(unname(REASON_LABELS[rtype]),
                         levels = unname(REASON_LABELS)))

n_unit      <- nrow(plant_festival)
n_unit_coded<- sum(!is.na(plant_festival$reason_types))
cat("\n=== 選定理由の解析単位 ===\n")
cat("祭り×植物分類群:", n_unit, "単位（うち理由が読み取れたもの", n_unit_coded, "）\n")
cat("1単位あたりの理由類型数:\n")
print(table(vapply(plant_festival$reason_types,
                   function(z) if (is.na(z)) 0L else length(strsplit(z, "\\|")[[1]]),
                   integer(1))))

# --- 17a: 理由類型の全体分布（素 vs 補正）-----------------------------------
# 「その理由を挙げた 祭り×植物 単位」の割合を、ウェイト付きでも計算する
u_reason <- table(factor(reason_long$festival, levels = fest_design$festival),
                  reason_long$rtype)
# ※ ここでは祭り単位ではなく単位（祭り×植物）を数えるので行列は使わず直接集計
reason_share <- reason_long %>%
  count(rtype, rlabel, name = "raw_n") %>%
  left_join(
    reason_long %>%
      group_by(rtype) %>%
      summarise(w_n = sum(w), .groups = "drop"),
    by = "rtype"
  ) %>%
  mutate(
    raw_share = raw_n / n_unit_coded,
    w_share   = w_n / sum(plant_festival$w[!is.na(plant_festival$reason_types)])
  ) %>%
  arrange(desc(w_share))

cat("\n=== 選定理由の類型別シェア（素 vs 府県ウェイト補正）===\n")
print(as.data.frame(reason_share %>%
  transmute(類型 = rlabel, 件数 = raw_n,
            素のシェア = round(raw_share, 3),
            補正シェア = round(w_share, 3))))

p17a <- reason_share %>%
  mutate(rlabel = fct_reorder(rlabel, w_share)) %>%
  select(rlabel, raw_share, w_share) %>%
  pivot_longer(-rlabel, names_to = "kind", values_to = "share") %>%
  mutate(kind = factor(kind, levels = c("raw_share", "w_share"),
                       labels = c("素のシェア（30祭りの標本）",
                                  "府県ウェイト補正後"))) %>%
  ggplot(aes(x = share, y = rlabel, fill = kind)) +
  geom_col(position = position_dodge(width = 0.72), width = 0.68, alpha = 0.9) +
  geom_text(aes(label = scales::percent(share, accuracy = 1)),
            position = position_dodge(width = 0.72), hjust = -0.15, size = 2.8,
            family = "HiraginoSans-W3") +
  scale_fill_manual(values = c("#BDBDBD", "#D62728"), name = NULL) +
  scale_x_continuous(labels = scales::percent, limits = c(0, 0.62)) +
  labs(
    title = "植物を選ぶ理由の類型",
    subtitle = paste0("解析単位＝祭り×植物分類群（n = ", n_unit_coded,
                      "、部位の重複は集約、理由は和集合）\n",
                      "1単位が複数類型を持つため合計は100%を超える"),
    x = "その理由を挙げた単位の割合", y = NULL
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "bottom",
        panel.grid.major.y = element_blank())

ggsave(file.path(OUTPUT_DIR, "17a_reason_types_overall.png"), p17a,
       width = 9, height = 6, dpi = 150)

# --- 17b: 主要植物ごとの理由構成（ウェイト付き）-----------------------------
major_taxa <- prev_tax %>% filter(raw_n >= 3) %>% pull(taxon)

reason_by_taxon <- reason_long %>%
  filter(resource_taxon %in% major_taxa) %>%
  group_by(resource_taxon, rtype, rlabel) %>%
  summarise(w_n = sum(w), .groups = "drop") %>%
  left_join(
    plant_festival %>%
      filter(resource_taxon %in% major_taxa, !is.na(reason_types)) %>%
      group_by(resource_taxon) %>%
      summarise(w_tot = sum(w), n_fes = n(), .groups = "drop"),
    by = "resource_taxon"
  ) %>%
  mutate(share = w_n / w_tot,
         taxon_label = paste0(resource_taxon, "（", n_fes, "祭り）"))

taxon_ord <- prev_tax %>% filter(raw_n >= 3) %>% arrange(w_prev) %>% pull(taxon)
reason_by_taxon <- reason_by_taxon %>%
  mutate(taxon_label = factor(taxon_label,
    levels = unique(taxon_label[order(match(resource_taxon, taxon_ord))])))

# 多重ラベルのため積み上げ棒だと合計が100%を超えて読みにくい。
# 「その植物を使う祭りのうち、その理由を挙げた割合」をヒートマップで示す。
reason_grid <- expand_grid(
  taxon_label = levels(reason_by_taxon$taxon_label),
  rlabel      = factor(unname(REASON_LABELS), levels = unname(REASON_LABELS))
) %>%
  left_join(reason_by_taxon %>% select(taxon_label, rlabel, share),
            by = c("taxon_label", "rlabel")) %>%
  mutate(share = ifelse(is.na(share), 0, share),
         taxon_label = factor(taxon_label, levels = levels(reason_by_taxon$taxon_label)))

p17b <- ggplot(reason_grid, aes(x = rlabel, y = taxon_label, fill = share)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = ifelse(share > 0, scales::percent(share, accuracy = 1), "")),
            size = 2.7, family = "HiraginoSans-W3",
            color = ifelse(reason_grid$share > 0.5, "white", "gray20")) +
  scale_fill_gradient(low = "#FFF7EC", high = "#B30000",
                      labels = scales::percent, name = "その理由を挙げた割合") +
  labs(
    title = "主要植物ごとの選定理由の構成",
    subtitle = paste0("3祭り以上で使われた植物。セル＝その植物を使う祭りのうち、",
                      "その理由が語られた割合（府県ウェイト補正後）。\n",
                      "1単位が複数類型を持つため行の合計は100%を超える。"),
    x = NULL, y = NULL
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        panel.grid = element_blank(),
        axis.text.x = element_text(angle = 30, hjust = 1),
        legend.position = "bottom")

ggsave(file.path(OUTPUT_DIR, "17b_reason_by_plant.png"), p17b,
       width = 10, height = 6.5, dpi = 150)

write.csv(
  plant_festival %>%
    select(pref, festival, resource_taxon, n_parts, parts, reason_types, w),
  file.path(OUTPUT_DIR, "plant_festival_units.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)
write.csv(
  reason_share %>% select(rtype, rlabel, raw_n, raw_share, w_n, w_share),
  file.path(OUTPUT_DIR, "reason_type_share.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

# ==============================================================================
# 図20: 利用方法カテゴリー × 代替可能性
#   松明内での植物の役割（燃焼材・構造材・化粧材・装飾材・結束材など）ごとに
#   代替可能性スコアの分布を積み上げ比率バーで示す。
#   「装飾・化粧材は代替不可が多い」等の仮説を検証。
# ==============================================================================

subst_colors_20 <- c(
  "代替可（1）"    = "#4DAF4A",
  "代替困難（2）"  = "#FF7F00",
  "代替不可（3）"  = "#E41A1C"
)

use_subst_df <- resource_df %>%
  filter(!is.na(use_types), !is.na(subst_score)) %>%
  mutate(use_split = str_split(use_types, "\\|")) %>%
  tidyr::unnest(use_split) %>%
  filter(use_split != "", use_split != "その他") %>%
  mutate(
    use_label  = factor(use_split,
                        levels = c("燃焼材","構造材","化粧材","装飾材","結束材","食材")),
    subst_label = factor(subst_score, levels = 1:3,
                         labels = c("代替可（1）", "代替困難（2）", "代替不可（3）"))
  ) %>%
  count(use_label, subst_label) %>%
  group_by(use_label) %>%
  mutate(total = sum(n), pct = n / total * 100) %>%
  ungroup() %>%
  mutate(use_label = fct_reorder(use_label,
                                  ifelse(subst_label == "代替不可（3）", pct, 0),
                                  .fun = sum))

p20_use_subst <- ggplot(use_subst_df,
       aes(x = use_label, y = pct, fill = subst_label)) +
  geom_col(position = "stack", width = 0.65) +
  geom_text(aes(label = ifelse(pct >= 8, paste0(round(pct), "%\n(n=", n, ")"), "")),
            position = position_stack(vjust = 0.5),
            size = 2.8, color = "white", fontface = "bold",
            family = "HiraginoSans-W3") +
  coord_flip() +
  scale_fill_manual(values = subst_colors_20, name = "代替可能性") +
  scale_y_continuous(labels = function(x) paste0(x, "%"),
                     expand = c(0, 0), limits = c(0, 102)) +
  labs(
    title    = "利用方法カテゴリー別 代替可能性の分布",
    subtitle = "バーは「代替不可（3）の割合」で降順ソート。バー内に割合と件数を表示。",
    x = "松明における利用方法",
    y = "割合（%）",
    caption = paste0("N = ", nrow(resource_df %>% filter(!is.na(use_types), !is.na(subst_score))),
                     " レコード（利用方法・代替可能性データ両方あり）")
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(
    plot.title  = element_text(face = "bold"),
    legend.position = "right",
    panel.grid.major.y = element_blank()
  )

ggsave(file.path(OUTPUT_DIR, "20_use_vs_substitutability.png"), p20_use_subst,
       width = 9, height = 5, dpi = 150)

# ==============================================================================
# 図21a: 日常利用スコアの全体分布
#   1=日常的に使う / 2=ほとんどない / 3=全くない の件数棒グラフ
# ==============================================================================

daily_label_lv <- c("1 日常的に使う", "2 ほとんどない", "3 全くない")
daily_colors   <- c("1 日常的に使う" = "#2CA02C",
                    "2 ほとんどない"  = "#FF7F0E",
                    "3 全くない"      = "#D62728")

daily_overall <- resource_df %>%
  filter(!is.na(daily_score)) %>%
  mutate(daily_label = factor(
    case_when(daily_score == 1 ~ "1 日常的に使う",
              daily_score == 2 ~ "2 ほとんどない",
              TRUE             ~ "3 全くない"),
    levels = daily_label_lv)) %>%
  count(daily_label)

p21a_daily_dist <- ggplot(daily_overall, aes(x = daily_label, y = n, fill = daily_label)) +
  geom_col(width = 0.6, show.legend = FALSE) +
  geom_text(aes(label = n), vjust = -0.4, size = 4,
            family = "HiraginoSans-W3") +
  scale_fill_manual(values = daily_colors) +
  scale_y_continuous(expand = expansion(mult = c(0, 0.15))) +
  labs(
    title    = "植物資源の日常利用スコア分布",
    subtitle = paste0("N = ", sum(daily_overall$n), " レコード（日常利用データあり）"),
    x = "日常利用スコア", y = "植物使用レコード数"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        panel.grid.major.x = element_blank())

ggsave(file.path(OUTPUT_DIR, "21a_daily_use_distribution.png"), p21a_daily_dist,
       width = 7, height = 5, dpi = 150)

# ==============================================================================
# 図21b: 日常利用 × 代替可能性（クロス集計・積み上げ比率バー）
# ==============================================================================

daily_subst <- resource_df %>%
  filter(!is.na(daily_score), !is.na(subst_score)) %>%
  mutate(
    daily_label = factor(
      case_when(daily_score == 1 ~ "1 日常的に使う",
                daily_score == 2 ~ "2 ほとんどない",
                TRUE             ~ "3 全くない"),
      levels = daily_label_lv),
    subst_label = factor(subst_score, levels = 1:3,
                         labels = c("代替可（1）", "代替困難（2）", "代替不可（3）"))
  ) %>%
  count(daily_label, subst_label) %>%
  group_by(daily_label) %>%
  mutate(pct = n / sum(n) * 100) %>%
  ungroup()

p21b_daily_subst <- ggplot(daily_subst,
       aes(x = daily_label, y = pct, fill = subst_label)) +
  geom_col(position = "stack", width = 0.6) +
  geom_text(aes(label = ifelse(pct >= 8, paste0(round(pct), "%"), "")),
            position = position_stack(vjust = 0.5),
            size = 3, color = "white", fontface = "bold",
            family = "HiraginoSans-W3") +
  scale_fill_manual(
    values = c("代替可（1）" = "#4DAF4A", "代替困難（2）" = "#FF7F00", "代替不可（3）" = "#E41A1C"),
    name = "代替可能性") +
  scale_y_continuous(labels = function(x) paste0(x, "%"),
                     expand = c(0, 0), limits = c(0, 102)) +
  labs(
    title    = "日常利用スコア別 代替可能性の分布",
    subtitle = "日常利用あり（1）の植物は代替可が多いか？",
    x = "日常利用スコア", y = "割合（%）"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "right",
        panel.grid.major.x = element_blank())

ggsave(file.path(OUTPUT_DIR, "21b_daily_vs_substitutability.png"), p21b_daily_subst,
       width = 8, height = 5, dpi = 150)

# ==============================================================================
# 図21c: 植物種別の日常利用スコア（頻出植物のみ）
# ==============================================================================

daily_by_resource <- resource_df %>%
  filter(!is.na(daily_score), resource_norm %in% top_resources) %>%
  group_by(resource_norm) %>%
  summarise(mean_daily = mean(daily_score, na.rm = TRUE),
            n = n(), .groups = "drop") %>%
  mutate(resource_norm = fct_reorder(resource_norm, mean_daily))

p21c_daily_resource <- ggplot(daily_by_resource,
       aes(x = resource_norm, y = mean_daily, fill = mean_daily)) +
  geom_col(width = 0.65) +
  geom_text(aes(label = paste0("n=", n)), hjust = -0.1, size = 2.8,
            family = "HiraginoSans-W3") +
  coord_flip() +
  scale_fill_gradient2(low = "#2CA02C", mid = "#FF7F0E", high = "#D62728",
                       midpoint = 2, limits = c(1, 3),
                       name = "平均スコア\n（1=有 3=無）") +
  scale_y_continuous(limits = c(0, 3.5),
                     breaks = 1:3,
                     labels = c("1\n日常的に使う", "2\nほとんどない", "3\n全くない")) +
  labs(
    title    = "植物種別 日常利用スコア（平均）",
    subtitle = "2祭り以上で使用された植物のみ。スコア低い（緑）ほど日常利用あり",
    x = NULL, y = "日常利用スコア（平均）"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "right")

ggsave(file.path(OUTPUT_DIR, "21c_daily_by_resource.png"), p21c_daily_resource,
       width = 9, height = 6, dpi = 150)

# ==============================================================================
# 図21d: 日常利用 × TEK利用タイプ（積み上げ比率バー）
# ==============================================================================

daily_tek <- resource_df %>%
  filter(!is.na(daily_score), !is.na(tek_types)) %>%
  mutate(
    daily_label = factor(
      case_when(daily_score == 1 ~ "1 日常的に使う",
                daily_score == 2 ~ "2 ほとんどない",
                TRUE             ~ "3 全くない"),
      levels = daily_label_lv),
    tek_split = str_split(tek_types, "\\|")
  ) %>%
  tidyr::unnest(tek_split) %>%
  filter(tek_split != "") %>%
  mutate(tek_label = recode(tek_split,
    "eco"  = "生態的特性", "aes"  = "美観・形状",
    "trad" = "伝統・慣習", "prag" = "実用・入手容易",
    "symb" = "象徴・宗教")) %>%
  count(tek_label, daily_label) %>%
  group_by(tek_label) %>%
  mutate(pct = n / sum(n) * 100) %>%
  ungroup() %>%
  mutate(tek_label = fct_reorder(tek_label,
                                  ifelse(daily_label == "3 全くない", pct, 0),
                                  .fun = sum))

p21d_daily_tek <- ggplot(daily_tek,
       aes(x = tek_label, y = pct, fill = daily_label)) +
  geom_col(position = "stack", width = 0.65) +
  geom_text(aes(label = ifelse(pct >= 8, paste0(round(pct), "%"), "")),
            position = position_stack(vjust = 0.5),
            size = 3, color = "white", fontface = "bold",
            family = "HiraginoSans-W3") +
  coord_flip() +
  scale_fill_manual(values = daily_colors, name = "日常利用") +
  scale_y_continuous(labels = function(x) paste0(x, "%"),
                     expand = c(0, 0), limits = c(0, 102)) +
  labs(
    title    = "TEK利用タイプ別 日常利用スコアの分布",
    subtitle = "バーは「日常利用なし（3）の割合」で降順ソート",
    x = "TEK利用タイプ", y = "割合（%）"
  ) +
  theme_bw(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold"),
        legend.position = "bottom",
        panel.grid.major.y = element_blank())

ggsave(file.path(OUTPUT_DIR, "21d_daily_vs_tek.png"), p21d_daily_tek,
       width = 9, height = 5, dpi = 150)

# ==============================================================================
# 図18: 生息地依存ネットワーク（祭り × 景観タイプ）
#   二部グラフ（bipartite）：祭り（左）― 景観タイプ（右）
#   ggplot で簡易的に可視化
# ==============================================================================

habitat_net <- resource_df %>%
  filter(!is.na(landscape_norm) & landscape_norm != "なし") %>%
  distinct(festival, landscape_norm) %>%
  group_by(landscape_norm) %>%
  mutate(n_festivals_using = n()) %>%
  ungroup()

# 景観タイプ別に依存祭り数を集計
habitat_summary <- habitat_net %>%
  group_by(landscape_norm) %>%
  summarise(
    n_festivals = n_distinct(festival),
    festivals   = paste(festival, collapse = "\n"),
    .groups = "drop"
  ) %>%
  arrange(desc(n_festivals))

cat("\n=== 生息地タイプ別依存祭り数 ===\n")
print(as.data.frame(habitat_summary))

# 二部グラフ用座標
hab_types  <- habitat_summary$landscape_norm
fest_using <- unique(habitat_net$festival)

node_df <- bind_rows(
  data.frame(
    name  = hab_types,
    type  = "habitat",
    x     = 1,
    y     = seq(1, length(hab_types)),
    stringsAsFactors = FALSE
  ),
  data.frame(
    name  = fest_using,
    type  = "festival",
    x     = 3,
    y     = seq(1, length(fest_using)),
    stringsAsFactors = FALSE
  )
)

edge_df <- habitat_net %>%
  left_join(node_df %>% filter(type == "habitat") %>%
              select(name, y_hab = y), by = c("landscape_norm" = "name")) %>%
  left_join(node_df %>% filter(type == "festival") %>%
              select(name, y_fest = y), by = c("festival" = "name"))

hab_colors_net <- c(
  "水田" = "#74C476", "湿地" = "#6BAED6",
  "森林" = "#31A354", "畑"   = "#FD8D3C",
  "荒地" = "#969696", "庭"   = "#BCBDDC"
)

p18_network <- ggplot() +
  geom_segment(data = edge_df,
               aes(x = 1, xend = 3, y = y_hab, yend = y_fest,
                   color = landscape_norm),
               alpha = 0.35, linewidth = 0.6) +
  geom_point(data = node_df %>% filter(type == "habitat"),
             aes(x = x, y = y, color = name), size = 8, alpha = 0.9) +
  geom_point(data = node_df %>% filter(type == "festival"),
             aes(x = x, y = y), size = 4, color = "gray40") +
  geom_text(data = node_df %>% filter(type == "habitat"),
            aes(x = x - 0.15, y = y, label = name),
            hjust = 1, size = 3.2, family = "HiraginoSans-W3") +
  geom_text(data = node_df %>% filter(type == "festival"),
            aes(x = x + 0.1, y = y, label = name),
            hjust = 0, size = 2.8, family = "HiraginoSans-W3") +
  scale_color_manual(values = hab_colors_net, guide = "none") +
  scale_x_continuous(limits = c(0.3, 4.5)) +
  annotate("text", x = 1, y = max(node_df$y) + 0.7,
           label = "生息地タイプ", fontface = "bold", size = 3.5,
           family = "HiraginoSans-W3") +
  annotate("text", x = 3, y = max(node_df$y) + 0.7,
           label = "祭り", fontface = "bold", size = 3.5,
           family = "HiraginoSans-W3") +
  labs(
    title = "生息地依存ネットワーク",
    subtitle = "線 = その生息地タイプの植物を使用する関係"
  ) +
  theme_void(base_family = "HiraginoSans-W3") +
  theme(plot.title = element_text(face = "bold", size = 13, hjust = 0.5),
        plot.subtitle = element_text(size = 9, hjust = 0.5, color = "gray40"))

ggsave(file.path(OUTPUT_DIR, "18_habitat_dependency_network.png"), p18_network,
       width = 11, height = 9, dpi = 150)

# 図19は削除

# ==============================================================================
# CSV出力
# ==============================================================================

write.csv(
  survey_df %>% select(festival, age_raw, participants_trend, tourists_trend,
                       belief, purpose, preservation, core_generation, challenges),
  file.path(OUTPUT_DIR, "survey_summary.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

write.csv(
  resource_df %>% select(festival, resource_raw, resource_norm,
                         landscape_norm, method_raw, subst_raw, subst_score,
                         embed_score, reason_raw, tek_types),
  file.path(OUTPUT_DIR, "resource_detail.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

write.csv(
  scale_df %>% select(festival, participants, tourists, n_resources, festival_type),
  file.path(OUTPUT_DIR, "scale_summary.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

write.csv(
  keystone_df %>% arrange(desc(keystone_score)),
  file.path(OUTPUT_DIR, "keystone_species.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

write.csv(
  embed_festival,
  file.path(OUTPUT_DIR, "embeddedness_by_festival.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

# 景観分類の判定結果（原文つき）— classify_landscape() の妥当性確認用
write.csv(
  resource_df %>%
    select(festival, resource_raw, landscape_raw, landscape_norm) %>%
    mutate(landscape_raw = str_replace_all(landscape_raw, "[\r\n]+", " ")) %>%
    arrange(landscape_norm, festival),
  file.path(OUTPUT_DIR, "landscape_mapping_check.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)

# ==============================================================================
# 完了
# ==============================================================================

cat("\n========================================\n")
cat("分析完了。出力 (", OUTPUT_DIR, "):\n", sep = "")
cat("  01〜14: 基本統計・組織・資源・信仰・目的 分析\n")
cat("  --- 生物文化多様性分析 ---\n")
cat("  15_cultural_keystone_species.png     文化的関鍵種\n")
cat("  16_procurement_embeddedness.png      文化-生態嵌入度\n")
cat("  17a_tek_types_overall.png            TEKタイプ全体分布\n")
cat("  17b_tek_by_festival.png              TEKタイプ祭り別構成\n")
cat("  18_habitat_dependency_network.png    生息地依存ネットワーク\n")
cat("  20_use_vs_substitutability.png      利用方法 × 代替可能性\n")
cat("  21a_daily_use_distribution.png     日常利用スコア全体分布\n")
cat("  21b_daily_vs_substitutability.png  日常利用 × 代替可能性\n")
cat("  21c_daily_by_resource.png          植物種別 日常利用スコア\n")
cat("  21d_daily_vs_tek.png               日常利用 × TEKタイプ\n")
cat("  10_scale_typology.png               2×2 規模・観光類型散布図\n")
cat("  --- CSV ---\n")
cat("  survey_summary.csv / resource_detail.csv / scale_summary.csv\n")
cat("  keystone_species.csv / embeddedness_by_festival.csv\n")
cat("========================================\n")
