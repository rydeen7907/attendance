"""
顔認証・勤怠登録 

Python 3.12.4

・imagesフォルダに顔画像を登録

・xlsxファイルのシートフォーマット 
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

"""


import cv2
import face_recognition
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import os

# 設定
EXCEL_FILE = '.xlsx' # フルパスで任意のファイル名
REGISTERED_FACES_SHEET = 'registered_faces'
ATTENDANCE_SHEET = 'attendance'
TOLERANCE = 0.6 # デフォルト値 ＝ 0.6 (顔認識時の許容範囲)

# ファイルの存在を確認する関数
def check_file_exists(file_path):
    """指定されたパスにファイルが存在するかどうかを確認します。

    Args:
        file_path (str): ファイルパス

    Returns:
        bool: ファイルが存在する場合はTrue、そうでない場合はFalse
    """
    if not os.path.isfile(file_path):
        messagebox.showerror("ファイルエラー", f"{file_path} が存在しません。適切なファイルを作成してください。")
        return False
    return True

# 既存の顔データを読み込む関数
def load_registered_faces(file_path, sheet_name):
    """Excelファイルから登録済みの顔データを読み込みます。

    Args:
        file_path (str): Excelファイルのパス
        sheet_name (str): 読み込むシート名

    Returns:
        tuple: 登録済みの顔データと顔エンコーディングのリスト
    """
    try:
        data = pd.read_excel(file_path, sheet_name=sheet_name)
        known_faces = []
        for _, row in data.iterrows():
            img_path = row['Image Path']
            if not os.path.exists(img_path):
                messagebox.showerror("ファイルエラー", f"画像ファイルが存在しません: {img_path}")
                raise FileNotFoundError(f"画像ファイルが存在しません: {img_path}")
            image = face_recognition.load_image_file(img_path)
            face_encoding = face_recognition.face_encodings(image)
            if face_encoding:
                known_faces.append(face_encoding[0])
            else:
                messagebox.showerror("顔認識エラー", f"顔が認識できませんでした: {img_path}")
                raise ValueError(f"顔が認識できませんでした: {img_path}")
        return data, known_faces
    except Exception as e:
        messagebox.showerror("データ読み込みエラー", str(e))
        raise e

# 勤怠データを更新する関数
def update_attendance(file_path, sheet_name, name, action):
    """Excelファイルに勤怠データを記録します。

    Args:
        file_path (str): Excelファイルのパス
        sheet_name (str): 書き込むシート名
        name (str): 従業員名
        action (str): 出勤または退勤
    """
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_entry = pd.DataFrame({'Name': [name], 'Timestamp': [timestamp], 'Action': [action]})
    
    try:
        attendance_data = pd.read_excel(file_path, sheet_name=sheet_name)
        attendance_data = pd.concat([attendance_data, new_entry], ignore_index=True)
    except FileNotFoundError:
        attendance_data = new_entry
    
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        attendance_data.to_excel(writer, sheet_name=sheet_name, index=False)

# 顔認証を行い、認識された従業員名を返す関数
def recognize_faces(known_faces, registered_faces_data, tolerance=TOLERANCE):
    """カメラを使用して顔認証を行い、認識された従業員名を返します。

    Args:
        known_faces (list): 登録済みの顔エンコーディングのリスト
        registered_faces_data (pd.DataFrame): 登録済みの顔データ
        tolerance (float, optional): 顔認識の許容範囲. Defaults to TOLERANCE.

    Returns:
        set: 認識された従業員名のセット
    """
    cap = cv2.VideoCapture(0)
    recognized_names = set()

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            messagebox.showerror("エラー", "カメラからフレームを取得できません")
            break

        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        names = []  # 認識された名前を格納するリスト
        for face_encoding in face_encodings:
            matches = face_recognition.compare_faces(known_faces, face_encoding, tolerance=tolerance)
            name = "Unknown" 
            if True in matches:
                first_match_index = matches.index(True)
                name = registered_faces_data.iloc[first_match_index]['Name']
                recognized_names.add(name)
            names.append(name)

        # 認証できた顔を表示 (デバッグ用)
        for (top, right, bottom, left), name in zip(face_locations, names):
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
            cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_DUPLEX, 0.5, (255, 255, 255), 1)

        cv2.imshow('Recognition', frame)
        if cv2.waitKey(1) & 0xFF == 27:  # 'ESC'キーで勤怠登録
            break

    cap.release()
    cv2.destroyAllWindows()
    return recognized_names

# GUIのボタン押下時の処理
def handle_attendance(action):
    """出勤/退勤ボタン押下時の処理を行います。

    Args:
        action (str): 'in' for 出勤, 'out' for 退勤
    """
    action_str = "出勤" if action == 'in' else "退勤"
    try:
        recognized_names = recognize_faces(known_faces, registered_faces_data)
        if recognized_names:
            messagebox.showinfo("認証成功", f"以下のユーザーが{action_str}しました：{', '.join(recognized_names)}")
            for name in recognized_names:
                update_attendance(EXCEL_FILE, ATTENDANCE_SHEET, name, action_str)
        else:
            messagebox.showwarning("認証失敗", "顔認証できませんでした。再試行してください。")
    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")

# GUIの構築
def create_gui():
    """勤怠管理システムのGUIを作成します。
    """
    root = tk.Tk()
    root.title("勤怠管理システム")

    label = tk.Label(root, text="出勤の場合は '出勤'ボタン、退勤の場合は '退勤'ボタンを押してください")
    label.pack(pady=10)

    button_in = tk.Button(root, text="出勤", command=lambda: handle_attendance('in'))
    button_in.pack(side=tk.LEFT, padx=20)

    button_out = tk.Button(root, text="退勤", command=lambda: handle_attendance('out'))
    button_out.pack(side=tk.RIGHT, padx=20)

    root.mainloop()

# メイン処理
if __name__ == "__main__":
    if not check_file_exists(EXCEL_FILE):
        exit(1)

    try:
        # 登録済みの顔データと顔エンコーディングの読み込み
        registered_faces_data, known_faces = load_registered_faces(EXCEL_FILE, REGISTERED_FACES_SHEET)
        # GUIの実行
        create_gui()
    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")