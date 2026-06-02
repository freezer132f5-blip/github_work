# ノートPC比較プロジェクト

## プロジェクト概要
ノートPCの購入検討のために、条件に合うモデルを調査・比較したExcel表を管理するプロジェクト。

## ファイル構成
- `ノートPC比較表.xlsx` - 比較表本体（Excel）
- `create_laptop_comparison.py` - Excelファイル生成スクリプト（openpyxl使用）

## 比較条件

### マスト条件
- Windows Hello対応IRカメラ搭載（顔認証）
- 税込17万円以下

### 任意条件
- CPU：i3相当以上
- メモリ：16GB
- SSD：256GB
- USB給電（USB-C Power Delivery対応）

### 除外条件
- 中国メーカー（Lenovo等）は対象外

## 比較対象モデル（2026年時点）
| メーカー | モデル | 税込価格 |
|---------|--------|---------|
| VAIO | SX14-R（2026年5月） | 159,800円〜（特別価格） |
| 富士通 | FMV Note A（2026年1月） | 169,800円〜 |
| NEC | LAVIE Direct N15（2026年春） | 要確認 |
| dynabook | Cシリーズ（2026年春） | 要確認 |
| HP | EliteBook 640 G11 | 140,000〜160,000円（推定） |

## Excel更新手順
1. `create_laptop_comparison.py` を編集してモデルデータを修正
2. スクリプトを実行：`python3 create_laptop_comparison.py`
3. GitHubにプッシュ

## GitHubリポジトリ
`https://github.com/freezer132f5-blip/github_work.git`

## 注意事項
- 価格・スペックは変動するため、購入前に各メーカー公式サイトで最新情報を確認すること
- 「要確認」のモデルは各公式直販サイトで価格確認が必要
