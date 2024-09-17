import tkinter as tk
import subprocess
import os
import serial
import threading
import time
from serial.tools import list_ports
from concurrent.futures import ThreadPoolExecutor
import numpy as np
import matplotlib.pyplot as plt
import folium
from folium.plugins import MarkerCluster
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import webbrowser
import winsound
import threading

# グローバル変数
ser = None
logging_active = True
executor = ThreadPoolExecutor(max_workers=3)
buffer = []
buffer_size = 20
file_path = None
root = None
graph_window = None
map_window = None
map_filename = None
enter_press_count = 0

# ロード画面表示
def show_loading_screen():
    global root
    root = tk.Tk()
    root.geometry("400x300")
    root.title("Loading...")

    loading_label = tk.Label(root, text="Loading, please wait...", font=('Helvetica', 20))
    loading_label.pack(expand=True)

    progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)
    progress_bar.pack(pady=20)

    root.after(100, lambda: load_main_screen(progress_bar))

    root.mainloop()

def load_main_screen(progress_bar):
    for i in range(100):
        progress_bar["value"] = i + 1
        root.update_idletasks()
        time.sleep(0.05)

    root.destroy()
    show_start_screen()

# スタート画面表示
def show_start_screen():
    global root
    root = tk.Tk()
    root.geometry("800x600")
    root.title("HPA_Pilot Application")

    main_frame = tk.Frame(root)
    main_frame.pack(expand=True, fill="both")

    start_label = tk.Label(main_frame, text="HPA_PilotApplication", font=('Helvetica', 40))
    start_label.pack(pady=50, anchor='center')

    input_frame = tk.Frame(main_frame)
    input_frame.pack(pady=20)

    save_name_label = tk.Label(input_frame, text="保存名:", font=('Helvetica', 20))
    save_name_label.grid(row=0, column=0, padx=10, pady=10)

    save_name_entry = tk.Entry(input_frame, font=('Helvetica', 20), width=30)
    save_name_entry.grid(row=0, column=1, padx=10, pady=10)

    start_button = tk.Button(main_frame, text="Start", font=('Helvetica', 30), command=lambda: start_application(save_name_entry.get()))
    start_button.pack(pady=20, anchor='center')

    credit_label = tk.Label(main_frame, text="@SkyProject_2022~2024", font=('Helvetica', 15))
    credit_label.pack(side='bottom', pady=10)

    # カウントダウンラベルを追加
    countdown_label = tk.Label(main_frame, font=('Helvetica', 20))
    countdown_label.pack(side='bottom', anchor='sw', padx=10, pady=10)

    # システム終了ボタンにカウントダウン機能を追加
    exit_button = tk.Button(main_frame, text="システム終了", font=('Helvetica', 15),
                            command=lambda: start_countdown(exit_button, countdown_label))
    exit_button.pack(side='bottom', anchor='sw', padx=10, pady=10)

    root.mainloop()

# カウントダウン機能
def start_countdown(button, countdown_label):
    button.config(state=tk.DISABLED)  # ボタンを無効化
    countdown_label.config(text="3秒後にシステムが終了します")
    countdown_time = 3

    def countdown():
        nonlocal countdown_time
        if countdown_time >= 0:
            countdown_label.config(text=f"システム終了まで {countdown_time} 秒")
            countdown_label.config(fg='red')  # ボタンの色を赤に変更
            countdown_time -= 1
            root.after(1000, countdown)  # 1秒後に再度この関数を呼び出す
        else:
            root.destroy()  # カウントダウンが0になったらシステムを終了する

    countdown()  # カウントダウンを開始

def start_application(filename):
    global file_path

    if not filename:
        filename = "data_log"
    filename = filename + ".txt"

    # 保存場所を指定するダイアログを表示
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=filename, title="保存場所を選択してください",
                                             filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
    if not file_path:
        messagebox.showwarning("保存キャンセル", "保存場所が選択されませんでした。保存がキャンセルされました。")
        return

    root.destroy()
    show_main_display(file_path)

# メイン画面表示
def show_main_display(filename):
    global root, ser, logging_active, enter_press_count

    root = tk.Tk()
    root.geometry("1300x800")
    enter_press_count = 0

    canvas = tk.Canvas(root, width=1300, height=800)
    canvas.pack()

    draw_lines(canvas)

    # ラベルの設定
    label_speed_text = tk.Label(root, text="速度:", font=('Helvetica', 40))
    label_speed_text.place(x=50, y=150)

    label_rpm_text = tk.Label(root, text="回転数:", font=('Helvetica', 40))
    label_rpm_text.place(x=670, y=150)

    label_altitude_text = tk.Label(root, text="高度:", font=('Helvetica', 40))
    label_altitude_text.place(x=50, y=500)

    label_speed_value = tk.Label(root, text="0.0", font=('Helvetica', 140))
    label_speed_value.place(x=195, y=100)

    label_rpm_value = tk.Label(root, text="0", font=('Helvetica', 140))
    label_rpm_value.place(x=865, y=100)

    label_altitude_value = tk.Label(root, text="0", font=('Helvetica', 140))
    label_altitude_value.place(x=195, y=450)

    label_speed_unit = tk.Label(root, text="m/s", font=('Helvetica', 40))
    label_speed_unit.place(x=555, y=315)

    label_rpm_unit = tk.Label(root, text="rpm", font=('Helvetica', 40))
    label_rpm_unit.place(x=1170, y=315)

    label_altitude_unit = tk.Label(root, text="cm", font=('Helvetica', 40))
    label_altitude_unit.place(x=555, y=625)

    # バーの設定
    bar_height = 280
    bar_width = 50
    bar_canvas = tk.Canvas(root, width=bar_width, height=bar_height, bg='white', highlightthickness=1, highlightbackground='black')
    bar_canvas.place(x=5, y=400)

    # SDカード状況のラベル
    label_sd_status = tk.Label(root, text="SD: NO", font=('Helvetica', 30))
    label_sd_status.place(x=680, y=420)

    # 接続ボタンと状態ラベル
    label_connection_status = tk.Label(root, text="No Connection", font=('Helvetica', 20))
    label_connection_status.place(x=800, y=540)

    com_label = tk.Label(root, text="COMポート:", font=('Helvetica', 20))
    com_label.place(x=680, y=480)

    available_ports = [port.device for port in list_ports.comports()]

    selected_com_port = tk.StringVar()
    selected_com_port.set(available_ports[0] if available_ports else "")

    com_menu = tk.OptionMenu(root, selected_com_port, *available_ports)
    com_menu.config(font=('Helvetica', 20))
    com_menu.place(x=840, y=475)

    connect_button = tk.Button(root, text="接続", font=('Helvetica', 20), command=lambda: connect_com(selected_com_port, label_connection_status, label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value))
    connect_button.place(x=680, y=530)

    root.bind('<Return>', lambda event: on_enter(canvas, bar_canvas, label_speed_text, label_rpm_text, label_altitude_text, label_speed_value, label_rpm_value, label_altitude_value, label_speed_unit, label_rpm_unit, label_altitude_unit, label_sd_status, label_connection_status, com_label, com_menu, connect_button))

    root.mainloop()

def draw_lines(canvas):
    width = 1300
    height = 800
    canvas.create_line(width / 2, 0, width / 2, height, fill="black")
    canvas.create_line(0, height / 2, width, height / 2, fill="black")

def connect_com(selected_com_port, label_connection_status, label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value):
    global ser
    com_port = selected_com_port.get()
    try:
        ser = serial.Serial(com_port, 115200)
        label_connection_status.config(text="Connected")
        start_data_update_thread(label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value)
    except serial.SerialException as e:
        label_connection_status.config(text="No Connection")
        print(f"Failed to open serial port: {e}")

def start_data_update_thread(label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value):
    executor.submit(update_data, label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value)

def update_data(label_sd_status, filename, bar_canvas, label_speed_value, label_rpm_value, label_altitude_value):
    global ser
    while ser is not None and ser.is_open:
        try:
            start_time = time.time()
            line = ser.readline().decode('utf-8').strip()
            check_sd_card(label_sd_status)
            save_data_to_sd(line, filename)
            data = line.split(',')

            speed_value = None
            rpm_value = None
            altitude_value = None

            for i in range(len(data)):
                if data[i] == "AIR":
                    speed_value = data[i + 1]
                elif data[i] == "RS":
                    rpm_value = data[i + 1]
                elif data[i] == "ALT":
                    altitude_value = data[i + 1]

            if speed_value is not None:
                label_speed_value.config(text=speed_value)
            if rpm_value is not None:
                label_rpm_value.config(text=rpm_value)
            if altitude_value is not None:
                label_altitude_value.config(text=altitude_value)
                update_alt_bar(bar_canvas, altitude_value)

            end_time = time.time()
            print(f"データ処理とバッファリングにかかった時間: {end_time - start_time:.6f}秒")

            time.sleep(0.025)

        except Exception as e:
            print(f"Error: {e}")

def check_sd_card(label_sd_status):
    if os.path.exists('D:/'):
        label_sd_status.config(text="SD: OK")
    else:
        label_sd_status.config(text="SD: NO")

def save_data_to_sd(data, filename):
    global logging_active, buffer
    if logging_active:
        buffer.append(data + '\n')
        if len(buffer) >= buffer_size:
            try:
                with open(filename, 'a') as file:
                    file.writelines(buffer)
                buffer.clear()
            except Exception as e:
                print(f"Error saving data to SD card: {e}")

def update_alt_bar(bar_canvas, alt_value):
    bar_height = 280
    bar_width = 50
    alt_percentage = max(0.0, min(float(alt_value) / 100.0, 1.0))
    fill_height = bar_height * alt_percentage
    bar_canvas.delete("all")

    color = "red" if float(alt_value) >= 100.0 else "green"
    bar_canvas.create_rectangle(0, 0, bar_width, bar_height, outline='black')
    bar_canvas.create_rectangle(0, bar_height - fill_height, bar_width, bar_height, fill=color)

# 音を鳴らす関数
def play_buzzer():
    winsound.Beep(2500, 3000)  # 2500Hzの音を3秒間鳴らす

# Enterキーが押された時の処理
def on_enter(canvas, bar_canvas, label_speed_text, label_rpm_text, label_altitude_text, label_speed_value, label_rpm_value, label_altitude_value, label_speed_unit, label_rpm_unit, label_altitude_unit, label_sd_status, label_connection_status, com_label, com_menu, connect_button):
    global enter_press_count, ser, logging_active

    enter_press_count += 1
    if enter_press_count == 1:
        hide_all_widgets(canvas, bar_canvas, label_speed_text, label_rpm_text, label_altitude_text, label_speed_value, label_rpm_value, label_altitude_value, label_speed_unit, label_rpm_unit, label_altitude_unit, label_sd_status, label_connection_status, com_label, com_menu, connect_button)
        threading.Thread(target=play_buzzer).start()  # 別スレッドで音を鳴らす
        canvas.create_rectangle(0, 0, 1300, 800, fill='red')
    elif enter_press_count == 2:
        logging_active = False
        if ser is not None and ser.is_open:
            ser.close()
        ser = None
        show_data_analysis()

def hide_all_widgets(canvas, bar_canvas, label_speed_text, label_rpm_text, label_altitude_text, label_speed_value, label_rpm_value, label_altitude_value, label_speed_unit, label_rpm_unit, label_altitude_unit, label_sd_status, label_connection_status, com_label, com_menu, connect_button):
    label_speed_text.place_forget()
    label_rpm_text.place_forget()
    label_altitude_text.place_forget()
    label_speed_value.place_forget()
    label_rpm_value.place_forget()
    label_altitude_value.place_forget()
    label_speed_unit.place_forget()
    label_rpm_unit.place_forget()
    label_altitude_unit.place_forget()
    label_sd_status.place_forget()
    label_connection_status.place_forget()
    com_label.place_forget()
    com_menu.place_forget()
    connect_button.place_forget()
    canvas.delete("all")
    bar_canvas.place_forget()

# データ解析と地図表示
def show_data_analysis():
    global root
    root.destroy()
    select_file()

def select_file():
    global file_path, root
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="データファイルを選択してください",
        filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
    )
    if file_path:
        root.deiconify()
        show_graph_window()
    else:
        print("ファイルが選択されませんでした。")

def show_graph_window():
    global graph_window, map_filename
    graph_window = tk.Toplevel()
    graph_window.title("グラフ表示")
    graph_window.geometry("1920x1080")

    fig, axs = plt.subplots(2, 2, figsize=(16, 10))

    first_date, first_time, rs_data, bno_data, alt_data, air_data, gps_data = process_file(file_path)
    max_length = max(len(rs_data), len(bno_data['x']), len(alt_data), len(air_data))
    rs_index = np.arange(len(rs_data))
    bno_index = np.arange(len(bno_data['x']))
    alt_index = np.arange(len(alt_data))
    air_index = np.arange(len(air_data))

    axs[0, 0].plot(rs_index, rs_data, label='RS (rpm)', color='blue')
    axs[0, 0].set_title('RS (rpm)')
    axs[0, 0].set_xlabel('Index')
    axs[0, 0].set_ylabel('RS (rpm)')
    axs[0, 0].grid(True)
    axs[0, 0].set_xlim([0, max_length])

    if bno_data['x'] and bno_data['y'] and bno_data['z']:
        axs[0, 1].plot(bno_index, bno_data['x'], label='BNO X', color='red')
        axs[0, 1].plot(bno_index, bno_data['y'], label='BNO Y', color='green')
        axs[0, 1].plot(bno_index, bno_data['z'], label='BNO Z', color='orange')
        axs[0, 1].set_title('BNO (Euler Angles)')
        axs[0, 1].set_xlabel('Index')
        axs[0, 1].set_ylabel('Angle (degrees)')
        axs[0, 1].legend()
        axs[0, 1].grid(True)
        axs[0, 1].set_xlim([0, max_length])

    if alt_data:
        axs[1, 0].plot(alt_index, alt_data, label='ALT (cm)', color='purple')
        axs[1, 0].set_title('ALT (cm)')
        axs[1, 0].set_xlabel('Index')
        axs[1, 0].set_ylabel('Altitude (cm)')
        axs[1, 0].grid(True)
        axs[1, 0].set_xlim([0, max_length])

    if air_data:
        axs[1, 1].plot(air_index, air_data, label='AIR (m/s)', color='brown')
        axs[1, 1].set_title('AIR (m/s)')
        axs[1, 1].set_xlabel('Index')
        axs[1, 1].set_ylabel('Speed (m/s)')
        axs[1, 1].grid(True)
        axs[1, 1].set_xlim([0, max_length])

    if first_date and first_time:
        plt.figtext(0.5, 0.98, f'Date: {first_date} Time: {first_time}', ha='center', fontsize=10)

    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=graph_window)
    canvas.draw()
    canvas.get_tk_widget().pack()

    show_map_button = tk.Button(graph_window, text="地図表示", font=('Helvetica', 10, 'bold'), fg='red', command=show_map_window)
    show_map_button.pack(pady=20)
    show_map_button.place(relx=0.51, rely=0.5, anchor='center')

    if map_filename:
        file_label = tk.Label(graph_window, text=f"地図ファイル: {os.path.basename(map_filename)}", font=('Helvetica', 10))
        file_label.pack(pady=10)
        file_label.place(relx=0.55, rely=0.5, anchor='w')

    quit_button = tk.Button(graph_window, text="終了", font=('Helvetica', 10, 'bold'), fg='red', command=on_quit)
    quit_button.pack(pady=20)
    quit_button.place(relx=0.51, rely=0.97, anchor='center')

def show_map_window():
    global map_filename
    try:
        if map_filename and os.path.exists(map_filename):
            webbrowser.open(f"file://{os.path.abspath(map_filename)}")
        else:
            messagebox.showerror("エラー", "地図ファイルが見つかりませんでした。")
    except Exception as e:
        print(f"Error in show_map_window: {e}")

def process_file(file_path, display_map=True):
    global map_filename
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()

        first_date = None
        first_time = None

        rs_data = []
        bno_data = {'x': [], 'y': [], 'z': []}
        alt_data = []
        air_data = []
        gps_data = []

        for i, line in enumerate(lines):
            if 'RS' in line and 'BNO' in line and 'ALT' in line and 'AIR' in line and 'GPS' in line:
                parts = line.split(',')
                try:
                    if i == 0:
                        first_date = parts[parts.index('GPS') + 1]
                        first_time = parts[parts.index('GPS') + 2]

                    rs_value = abs(float(parts[parts.index('RS') + 1]))
                    rs_data.append(rs_value)

                    bno_index = parts.index('BNO')
                    bno_data['x'].append(float(parts[bno_index + 1]))
                    bno_data['y'].append(float(parts[bno_index + 2]))
                    bno_data['z'].append(float(parts[bno_index + 3]))

                    alt_value = float(parts[parts.index('ALT') + 1])
                    alt_data.append(alt_value)

                    air_value = float(parts[parts.index('AIR') + 1])
                    air_data.append(air_value)

                    gps_index = parts.index('GPS')
                    lat = float(parts[gps_index + 3])
                    long = float(parts[gps_index + 4])
                    gps_data.append((lat, long))

                except (ValueError, IndexError) as e:
                    continue

        if display_map and gps_data:
            m = folium.Map(location=gps_data[0], zoom_start=15)
            marker_cluster = MarkerCluster().add_to(m)

            for lat, long in gps_data:
                folium.Marker(location=[lat, long]).add_to(marker_cluster)

            map_filename = 'map_with_gps.html'
            if os.path.exists(map_filename):
                base, ext = os.path.splitext(map_filename)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                map_filename = f"{base}_{counter}{ext}"

            m.save(map_filename)
            print(f"GPSデータをプロットした地図を '{map_filename}' に保存しました。")
        else:
            map_filename = None  # 地図が生成されなかった場合にNoneを設定
        return first_date, first_time, rs_data, bno_data, alt_data, air_data, gps_data
    except Exception as e:
        print(f"Error in process_file: {e}")
        return None, None, [], {'x': [], 'y': [], 'z': []}, [], [], []

def on_quit():
    try:
        # 現在のウィンドウを閉じて最初の画面に戻る
        graph_window.destroy()  # グラフ表示画面を閉じる
        show_start_screen()  # 最初の画面を再表示する
    except Exception as e:
        print(f"Error in on_quit: {e}")

# プログラムの開始
show_loading_screen()  # HPA_1
