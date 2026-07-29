# -*- coding: utf-8 -*-
import json
import os
import shutil
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import zipfile

APP_NAME = "항공자동조회"
APP_DIR_NAME = "AirAutoLookup"
DEFAULT_INSTALL_DIR = Path(os.environ.get("LOCALAPPDATA", r"C:\Users\Public")) / APP_DIR_NAME


def get_resource_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).resolve().parent / relative_path


def create_shortcut(target_exe: Path, shortcut_path: Path, icon_path: Path | None = None) -> None:
    try:
        import win32com.client

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.TargetPath = str(target_exe)
        shortcut.WorkingDirectory = str(target_exe.parent)
        if icon_path and icon_path.exists():
            shortcut.IconLocation = str(icon_path)
        shortcut.save()
    except Exception as exc:
        print(f"바로가기 생성 실패 ({shortcut_path}): {exc}")


def ensure_empty_hotels_manifest(install_dir: Path) -> Path:
    manifest_path = install_dir / "hotels-manifest.json"
    if not manifest_path.exists():
        empty_manifest = {
            "schema": "air-auto-lookup-flight-masters-v1",
            "updatedAt": "2026-07-29T00:00:00",
            "flightMasters": [],
        }
        manifest_path.write_text(json.dumps(empty_manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest_path


def extract_payload_member(zip_ref: zipfile.ZipFile, member: zipfile.ZipInfo, install_dir: Path) -> None:
    """실행 중인 이전 앱이 종료될 때까지 잠긴 파일 교체를 재시도한다."""
    deadline = time.monotonic() + 30
    while True:
        try:
            zip_ref.extract(member, install_dir)
            return
        except PermissionError as exc:
            if time.monotonic() >= deadline:
                raise RuntimeError(
                    "실행 중인 항공자동조회가 종료되지 않아 업데이트할 수 없습니다. "
                    "프로그램을 모두 닫은 뒤 다시 설치해 주세요."
                ) from exc
            time.sleep(0.5)


class InstallerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} 설치")
        self.root.geometry("520x360")
        self.root.resizable(False, False)

        self.install_dir_var = tk.StringVar(value=str(DEFAULT_INSTALL_DIR))
        self.create_desktop_shortcut_var = tk.BooleanVar(value=True)
        self.create_start_menu_shortcut_var = tk.BooleanVar(value=True)
        self.run_after_install_var = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value="설치 버튼을 누르면 항공자동조회가 컴퓨터에 설치됩니다.")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding=16)
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_label = ttk.Label(
            main_frame,
            text=f"✈ {APP_NAME} 설치 프로그램",
            font=("맑은 고딕", 14, "bold"),
        )
        header_label.pack(anchor="w", pady=(0, 12))

        dir_frame = ttk.LabelFrame(main_frame, text="설치 경로 선택", padding=10)
        dir_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(dir_frame, textvariable=self.install_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(dir_frame, text="찾아보기...", command=self.browse_dir).pack(side=tk.RIGHT)

        options_frame = ttk.LabelFrame(main_frame, text="설치 옵션", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(options_frame, text="바탕화면에 바로가기 만들기", variable=self.create_desktop_shortcut_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(options_frame, text="시작 메뉴에 바로가기 만들기", variable=self.create_start_menu_shortcut_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(options_frame, text="설치 완료 후 프로그램 실행", variable=self.run_after_install_var).pack(anchor="w", pady=2)

        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(6, 12))

        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)
        ttk.Label(progress_frame, textvariable=self.status_var, foreground="#444444").pack(anchor="w", pady=(4, 0))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.install_btn = ttk.Button(btn_frame, text="설치 시작", command=self.start_install)
        self.install_btn.pack(side=tk.RIGHT, padx=(6, 0))

        cancel_btn = ttk.Button(btn_frame, text="취소", command=self.root.destroy)
        cancel_btn.pack(side=tk.RIGHT)

    def browse_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.install_dir_var.get())
        if chosen:
            self.install_dir_var.set(chosen)

    def start_install(self) -> None:
        install_dir = Path(self.install_dir_var.get()).resolve()
        self.install_btn.configure(state=tk.DISABLED)
        self.status_var.set("설치 압축 해제 준비 중...")
        self.root.update_idletasks()

        try:
            install_dir.mkdir(parents=True, exist_ok=True)
            payload_zip = get_resource_path("payload.zip")

            if not payload_zip.exists():
                raise FileNotFoundError(f"번들 압축 파일(payload.zip)을 찾을 수 없습니다: {payload_zip}")

            self.status_var.set("파일 복사 및 설치 중...")
            self.progress_var.set(30)
            self.root.update_idletasks()

            with zipfile.ZipFile(payload_zip, "r") as zip_ref:
                for member in zip_ref.infolist():
                    target_file = install_dir / member.filename
                    if member.filename == "hotels-manifest.json" or member.filename.endswith("/hotels-manifest.json"):
                        if target_file.exists():
                            continue
                    extract_payload_member(zip_ref, member, install_dir)

            self.progress_var.set(70)
            self.status_var.set("노선 매니페스트 및 환경 설정 중...")
            self.root.update_idletasks()

            ensure_empty_hotels_manifest(install_dir)

            target_exe = install_dir / "항공자동조회.exe"
            icon_path = install_dir / "assets" / "air_auto_lookup_icon.ico"

            if self.create_desktop_shortcut_var.get():
                desktop = Path(os.environ.get("USERPROFILE", r"C:\Users\Public")) / "Desktop"
                shortcut = desktop / f"{APP_NAME}.lnk"
                create_shortcut(target_exe, shortcut, icon_path)

            if self.create_start_menu_shortcut_var.get():
                start_menu = Path(os.environ.get("APPDATA", r"C:\Users\Public")) / "Microsoft\Windows\Start Menu\Programs"
                shortcut = start_menu / f"{APP_NAME}.lnk"
                create_shortcut(target_exe, shortcut, icon_path)

            self.progress_var.set(100)
            self.status_var.set("설치가 성공적으로 완료되었습니다!")
            messagebox.showinfo("설치 완료", f"{APP_NAME} 설치가 완료되었습니다.\n설치 경로: {install_dir}")

            if self.run_after_install_var.get() and target_exe.exists():
                os.startfile(target_exe)

            self.root.destroy()
        except Exception as exc:
            self.install_btn.configure(state=tk.NORMAL)
            self.status_var.set("설치 실패")
            messagebox.showerror("설치 오류", f"설치 중 오류가 발생했습니다:\n{exc}")


def main() -> None:
    root = tk.Tk()
    app = InstallerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
