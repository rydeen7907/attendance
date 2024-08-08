顔認証による勤怠登録

python3.12.4で作成・動作確認済み

＝＝＝＝＝ 構成 ＝＝＝＝＝

images：顔画像登録用フォルダ(要作成)

ui_attendance.py：プログラム本体

kintai.xlsx：勤怠登録xlsxファイル(名前は任意)

シートフォーマット
sheet1 : registered_faces
| Image Path           | Name        |
|----------------------|-------------|
| images/employee1.jpg | Employee1   |
| images/employee2.jpg | Employee2   |

sheet2 : attendance
| Name      | Timestamp           | Action |
|-----------|---------------------|--------|
| Employee1 | 2023-10-01 08:00:00 | 出勤   |
| Employee2 | 2023-10-01 17:00:00 | 退勤   |

GUIで出退勤を選択しESCキーを押すと
従業員名と時間がxlsxファイルに反映される

プログラムの終了はGUIのバツ印をクリック
