import os
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
import tkinter as tk
from tkinter import ttk, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk, ImageGrab
import io
import win32clipboard
from io import BytesIO
import time
import pyautogui
import keyboard

class ImageViewer:
    def __init__(self, image):
        self.window = tk.Toplevel()
        self.window.title("분할된 이미지")
        self.window.geometry("1200x800")
        
        self.image_list = []
        
        self.main_frame = ttk.Frame(self.window)
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        ttk.Label(
            self.scrollable_frame,
            text="창을 클릭하면 모든 이미지가 순서대로 복사됩니다",
            font=('Arial', 12)
        ).pack(pady=10)

        self.grid_frame = ttk.Frame(self.scrollable_frame)
        self.grid_frame.pack(padx=20, pady=10)

        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.window.bind("<Button-1>", self.start_auto_copy)
        
        self.process_and_display_images(image)

    def send_to_clipboard(self, clip_type, data):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()

    def copy_to_clipboard(self, image):
        try:
            output = BytesIO()
            image.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]
            output.close()
            self.send_to_clipboard(win32clipboard.CF_DIB, data)
            return True
        except Exception as e:
            messagebox.showerror("오류", f"복사 중 오류 발생: {str(e)}")
            return False

    def start_auto_copy(self, event=None):
        if not self.image_list:
            messagebox.showwarning("경고", "복사할 이미지가 없습니다.")
            return

        response = messagebox.askquestion("자동 복사", 
            "모든 이미지를 순서대로 복사합니다.\n" +
            "1. 엑셀 창을 준비해주세요\n" +
            "2. 첫 번째 셀을 선택해주세요\n" +
            "3. '예'를 누르면 3초 후 시작됩니다\n" +
            "4. ESC를 누르면 중지됩니다")
        
        if response == 'yes':
            self.window.iconify()
            time.sleep(3)
            
            for i, image in enumerate(self.image_list):
                if keyboard.is_pressed('esc'):
                    messagebox.showinfo("중지", "작업이 중지되었습니다.")
                    break
                
                if self.copy_to_clipboard(image):
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.5)
                    
                    if (i + 1) % 3 == 0:
                        pyautogui.press('down')
                        pyautogui.press('left')
                        pyautogui.press('left')
                    else:
                        pyautogui.press('right')
                    
                    time.sleep(0.5)

            messagebox.showinfo("완료", "모든 이미지 복사가 완료되었습니다!")
            self.window.deiconify()

    def process_and_display_images(self, image):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
        
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        valid_contours = []
        min_area = 1000
        
        for contour in contours:
            area = cv2.contourArea(contour)
            if area > min_area:
                x, y, w, h = cv2.boundingRect(contour)
                if w > 50 and h > 50:
                    valid_contours.append((x, y, w, h))
        
        valid_contours.sort(key=lambda x: (x[1] // 100, x[0]))
        
        for i, (x, y, w, h) in enumerate(valid_contours):
            img_frame = ttk.Frame(self.grid_frame)
            img_frame.grid(row=i//3, column=i%3, padx=10, pady=10)
            
            margin = 2
            roi = image[y+margin:y+h-margin, x+margin:x+w-margin]
            
            if roi.size > 0:
                roi_rgb = cv2.cvtColor(roi, cv2.COLOR_BGR2RGB)
                pil_image = Image.fromarray(roi_rgb)
                
                display_size = (350, 250)
                display_image = pil_image.copy()
                display_image.thumbnail(display_size)
                
                self.image_list.append(pil_image)
                
                photo = ImageTk.PhotoImage(display_image)
                label = ttk.Label(img_frame, image=photo)
                label.image = photo
                label.pack()
                
                ttk.Label(img_frame, text=f"이미지 {i+1}", font=('Arial', 10)).pack(pady=5)

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("이미지 분할 프로그램")
        self.root.geometry("1000x800")
        
        self.image_list = []
        self.current_image = None
        self.preview_photos = []
        self.root.bind('<Control-v>', self.paste_image)
        
        self.create_widgets()
        
    def create_widgets(self):
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 왼쪽 프레임
        left_frame = ttk.Frame(self.main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(
            left_frame,
            text="Ctrl+V를 눌러 이미지를 붙여넣으세요",
            font=('Arial', 12)
        ).pack(pady=5)
        
        self.preview_label = ttk.Label(left_frame, text="현재 이미지 미리보기")
        self.preview_label.pack(pady=5)
        
        self.image_label = ttk.Label(left_frame)
        self.image_label.pack(pady=10)
        
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(
            button_frame,
            text="이미지 붙여넣기 (Ctrl+V)",
            command=lambda: self.paste_image(None)
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="이미지 추가",
            command=self.add_image
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="이미지 분할",
            command=self.process_images
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="초기화",
            command=self.reset_images
        ).pack(side=tk.LEFT, padx=5)
        
        # 오른쪽 프레임
        right_frame = ttk.Frame(self.main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)
        
        ttk.Label(
            right_frame,
            text="추가된 이미지 목록",
            font=('Arial', 12)
        ).pack(pady=5)
        
        self.image_count_label = ttk.Label(right_frame, text="(0개)")
        self.image_count_label.pack(pady=5)
        
        # 스크롤 가능한 미리보기 영역
        preview_canvas = tk.Canvas(right_frame)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=preview_canvas.yview)
        self.preview_frame = ttk.Frame(preview_canvas)
        
        self.preview_frame.bind(
            "<Configure>",
            lambda e: preview_canvas.configure(scrollregion=preview_canvas.bbox("all"))
        )
        
        preview_canvas.create_window((0, 0), window=self.preview_frame, anchor="nw")
        preview_canvas.configure(yscrollcommand=scrollbar.set)
        
        preview_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def paste_image(self, event=None):
        try:
            image = ImageGrab.grabclipboard()
            if image is None:
                messagebox.showwarning("경고", "클립보드에 이미지가 없습니다!")
                return
            
            self.current_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            preview = image.copy()
            preview.thumbnail((400, 400))
            photo = ImageTk.PhotoImage(preview)
            self.image_label.config(image=photo)
            self.image_label.image = photo
            
        except Exception as e:
            messagebox.showerror("오류", f"이미지 붙여넣기 실패: {str(e)}")
    
    def add_image(self):
        if self.current_image is None:
            messagebox.showwarning("경고", "먼저 이미지를 붙여넣어 주세요.")
            return
        
        self.image_list.append(self.current_image.copy())
        self.update_preview_list()
        messagebox.showinfo("성공", "이미지가 추가되었습니다.")
    
    def update_preview_list(self):
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        
        self.preview_photos.clear()
        self.image_count_label.config(text=f"({len(self.image_list)}개)")
        
        for i, img in enumerate(self.image_list):
            frame = ttk.Frame(self.preview_frame)
            frame.pack(pady=5, padx=5, fill=tk.X)
            
            pil_img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
            pil_img.thumbnail((200, 200))
            photo = ImageTk.PhotoImage(pil_img)
            self.preview_photos.append(photo)
            
            ttk.Label(frame, image=photo).pack(side=tk.LEFT, padx=5)
            ttk.Label(frame, text=f"이미지 {i+1}").pack(side=tk.LEFT, padx=5)
    
    def reset_images(self):
        if not self.image_list:
            messagebox.showinfo("알림", "초기화할 이미지가 없습니다.")
            return
        
        if messagebox.askyesno("확인", "모든 이미지를 초기화하시겠습니까?"):
            self.image_list.clear()
            self.update_preview_list()
            messagebox.showinfo("완료", "모든 이미지가 초기화되었습니다.")
    
    def process_images(self):
        if not self.image_list:
            messagebox.showwarning("경고", "추가된 이미지가 없습니다.")
            return
        
        try:
            total_height = sum(img.shape[0] for img in self.image_list)
            max_width = max(img.shape[1] for img in self.image_list)
            
            combined_image = np.zeros((total_height, max_width, 3), dtype=np.uint8)
            y_offset = 0
            
            for img in self.image_list:
                h, w = img.shape[:2]
                combined_image[y_offset:y_offset+h, :w] = img
                y_offset += h
            
            ImageViewer(combined_image)
            
            self.image_list.clear()
            self.update_preview_list()
            
        except Exception as e:
            messagebox.showerror("오류", f"이미지 처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
