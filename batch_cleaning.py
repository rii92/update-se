#!/usr/bin/env python3
"""
Batch Automation Cleaning FASIH-SM SE2026 (Standalone Engine)
Menjalankan otomatisasi kuesioner FASIH-SM dari file Excel (source/cleaning1.xlsx)
hanya dengan 1 perintah:
1. Otomatis membuka Google Chrome dengan profil terisolasi (./chrome_data).
2. Otomatis login SSO BPS jika belum login / sesi habis.
3. Otomatis melewati (filter) baris yang sudah SUCCESS sebelumnya.
4. Otomatis mengisi komentar Nomor KK (#0), mencentang Anomali Usaha & Keluarga,
   mengisi penjelasan anomali admin, dan submit kuesioner.
5. Otomatis merekap hasil ke file Excel (source/cleaning1_result.xlsx).
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from playwright.sync_api import BrowserContext, Page, sync_playwright
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

import config

console = Console()

PROGRESS_FILE = Path("source/cleaning1_progress.json")
RESULT_EXCEL_FILE = Path("source/cleaning1_result.xlsx")


def load_progress() -> Dict[str, dict]:
    """Memuat data checkpoint dari file json dan file Excel hasil (jika ada)."""
    progress = {}

    # 1. Baca dari JSON jika ada
    if PROGRESS_FILE.exists():
        try:
            with open(PROGRESS_FILE, "r") as f:
                progress = json.load(f)
        except Exception:
            pass

    # 2. Sinkronkan dari Excel result jika ada data yang belum masuk ke JSON
    if RESULT_EXCEL_FILE.exists():
        try:
            wb = openpyxl.load_workbook(RESULT_EXCEL_FILE, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                idx_str = str(row[0]).strip()
                status = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                if idx_str not in progress and status:
                    progress[idx_str] = {
                        "link": str(row[1]).strip() if len(row) > 1 and row[1] else "",
                        "komentar": str(row[2]).strip() if len(row) > 2 and row[2] else "#0",
                        "status": status,
                        "detail": str(row[4]).strip() if len(row) > 4 and row[4] else "",
                        "timestamp": str(row[5]).strip() if len(row) > 5 and row[5] else "",
                    }
        except Exception:
            pass

    return progress


def save_progress(data: Dict[str, dict]):
    """Menyimpan data checkpoint ke file json."""
    PROGRESS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(data, f, indent=2)


def export_result_excel(tasks: List[dict], progress_data: Dict[str, dict], output_path: Path):
    """Mengekspor rekapitulasi status eksekusi ke Excel."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hasil Cleaning"

    headers = ["No", "Link", "Komentar KK", "Status", "Detail", "Waktu Eksekusi"]
    ws.append(headers)

    for item in tasks:
        idx = item["index"]
        link = item["link"]
        rec = progress_data.get(str(idx), {})
        status = rec.get("status", "PENDING")
        detail = rec.get("detail", "")
        ts = rec.get("timestamp", "")
        ws.append([idx, link, item["komentar"], status, detail, ts])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def read_excel_source(file_path: Path) -> List[dict]:
    """Membaca daftar assignment dari Excel source."""
    if not file_path.exists():
        console.print(f"[bold red]❌ File sumber tidak ditemukan: {file_path}[/bold red]")
        sys.exit(1)

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    tasks = []
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        if not row or not row[0]:
            continue
        link = str(row[0]).strip()
        komentar = str(row[1]).strip() if len(row) > 1 and row[1] is not None else "#0"
        tasks.append({
            "index": row_idx,
            "link": link,
            "komentar": komentar
        })

    return tasks


def is_page_logged_out(page: Page) -> bool:
    """Mengecek apakah halaman saat ini berada pada status logout atau login SSO."""
    try:
        url = page.url.lower()
        if "sso.bps.go.id" in url or "/login" in url or "/auth/login" in url or "/oauth2/authorization" in url:
            return True
        if page.locator("#kc-form-login, #username, #password, .login-button").count() > 0:
            return True
        text = page.locator("body").inner_text() if page.locator("body").count() > 0 else ""
        if "Selamat Datang Kembali" in text or "Lanjutkan dengan SSO" in text or "Login SSO BPS" in text:
            return True
        return False
    except Exception:
        return False


def perform_sso_login(page: Page, username: str = config.SSO_USERNAME, password: str = config.SSO_PASSWORD) -> bool:
    """Melakukan login otomatis SSO BPS pada halaman browser."""
    console.print(f"[bold yellow]🔐 Login otomatis SSO BPS (User: {username})...[/bold yellow]")
    try:
        if "sso.bps.go.id" not in page.url:
            sso_btn = page.locator(
                "a:has-text('Lanjutkan dengan SSO'), a:has-text('Login SSO BPS'), a[href*='oauth2/authorization'], .login-button"
            )
            if sso_btn.count() > 0 and sso_btn.first.is_visible():
                sso_btn.first.click()
            else:
                page.goto("https://fasih-sm.bps.go.id/oauth2/authorization/ics", wait_until="domcontentloaded")
            page.wait_for_url("**/sso.bps.go.id/**", timeout=20000)

        page.wait_for_selector("#username", timeout=15000)
        page.fill("#username", username)
        page.fill("#password", password)
        page.click("#kc-login")

        page.wait_for_url("**/fasih-sm.bps.go.id/app/**", timeout=25000)
        console.print(f"[bold green]✔ Berhasil login SSO BPS sebagai {username}![/bold green]")
        time.sleep(2.0)
        return True
    except Exception as e:
        console.print(f"[bold red]❌ Gagal login otomatis SSO: {e}[/bold red]")
        return False


def process_single_assignment(
    page: Page,
    link: str,
    komentar_kk: str = "#0",
    anomaly_explanation: str = "Anomali diselesaikan oleh admin",
    retry_on_logout: bool = True,
) -> Tuple[bool, str]:
    """
    Menjalankan alur pembersihan kuesioner secara otomatis untuk satu assignment.
    """
    edit_url = link.strip()
    if not edit_url.endswith("/edit"):
        edit_url = edit_url.rstrip("/") + "/edit"

    try:
        page.goto(edit_url, wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        return False, f"Timeout membuka URL: {e}"

    time.sleep(1.0)

    # Cek jika sesi logout
    if is_page_logged_out(page):
        if retry_on_logout and perform_sso_login(page):
            return process_single_assignment(page, link, komentar_kk, anomaly_explanation, retry_on_logout=False)
        return False, "Sesi login habis / terlogout"

    # Tunggu aside menu kuesioner siap
    try:
        page.wait_for_selector("aside", timeout=20000)
    except Exception:
        if is_page_logged_out(page) and retry_on_logout:
            if perform_sso_login(page):
                return process_single_assignment(page, link, komentar_kk, anomaly_explanation, retry_on_logout=False)
        return False, "Menu sidebar (aside) kuesioner tidak muncul"

    time.sleep(1.0)

    # 1. Masuk ke blok SE2026 - P
    click_p_ok = page.evaluate("""(() => {
        const aside = document.querySelector("aside");
        if (!aside) return false;
        const sidebars = Array.from(aside.querySelectorAll("*")).filter(el => 
            el.children.length === 0 && (el.textContent.trim() === "SE2026 - P" || el.textContent.includes("SE2026 - P"))
        );
        if (sidebars.length === 0) return false;
        const target = sidebars[0];
        let clickable = target;
        while (clickable && clickable !== aside && !clickable.onclick && !clickable.className.includes("cursor-pointer")) {
            clickable = clickable.parentElement;
        }
        (clickable || target).click();
        return true;
    })()""")

    if not click_p_ok:
        return False, "Blok SE2026 - P tidak ditemukan di sidebar"

    time.sleep(1.5)

    # 2. Temukan pertanyaan Nomor KK & buka pop-up catatan
    click_note_ok = page.evaluate("""(() => {
        const rows = Array.from(document.querySelectorAll("div.tw\\\\:flex.tw\\\\:gap-2.tw\\\\:rounded-lg"));
        const kkRow = rows.find(r => r.innerText.includes("Nomor KK"));
        if (!kkRow) return "ERR: Row Nomor KK tidak ditemukan";
        const btns = Array.from(kkRow.querySelectorAll("button[title='Catatan'], button[aria-label='Catatan']"));
        const visibleBtn = btns.find(b => b.offsetWidth > 0 && b.offsetHeight > 0) || btns[0];
        if (!visibleBtn) return "ERR: Tombol Catatan Nomor KK tidak ditemukan";
        visibleBtn.scrollIntoView({ behavior: "smooth", block: "center" });
        visibleBtn.click();
        return "OK";
    })()""")

    if click_note_ok != "OK":
        return False, click_note_ok

    try:
        page.wait_for_selector("textarea[placeholder*='catatan' i]", timeout=8000)
    except Exception:
        return False, "Dialog catatan tidak muncul"

    # 3. Isi komentar Nomor KK dan simpan
    page.evaluate(f"""(() => {{
        const ta = document.querySelector("textarea[placeholder*='catatan' i]");
        if (!ta) return "ERR";
        const dialog = ta.closest("[role='dialog']") || document.querySelector("[class*='max-w-lg']");
        const valToFill = "{komentar_kk}";
        if (dialog && dialog.innerText.includes(valToFill)) {{
            const closeBtn = dialog.querySelector("button[aria-label='Dismiss']") || 
                             Array.from(dialog.querySelectorAll("button")).find(b => b.innerText.trim() === "Close");
            if (closeBtn) closeBtn.click();
            return "ALREADY";
        }}
        ta.focus();
        const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
        nativeSetter.call(ta, valToFill);
        ta.dispatchEvent(new Event("input", {{ bubbles: true }}));
        ta.dispatchEvent(new Event("change", {{ bubbles: true }}));

        const saveBtn = document.querySelector("button[title='Simpan']");
        if (saveBtn) saveBtn.click();

        setTimeout(() => {{
            const closeBtn = document.querySelector("button[aria-label='Dismiss']") || 
                             Array.from(document.querySelectorAll("button")).find(b => b.innerText.trim() === "Close");
            if (closeBtn) closeBtn.click();
        }}, 500);
        return "SAVED";
    }})()""")

    time.sleep(1.5)

    # 4. Masuk ke blok CATATAN
    click_catatan_ok = page.evaluate("""(() => {
        const aside = document.querySelector("aside");
        if (!aside) return false;
        const elements = Array.from(aside.querySelectorAll("*")).filter(el => el.children.length === 0 && el.textContent.trim() === "CATATAN");
        if (elements.length === 0) return false;
        const target = elements[0];
        let clickable = target;
        while (clickable && clickable !== aside && !clickable.onclick && !clickable.className.includes("cursor-pointer")) {
            clickable = clickable.parentElement;
        }
        (clickable || target).click();
        return true;
    })()""")

    if not click_catatan_ok:
        return False, "Menu CATATAN tidak ditemukan di sidebar"

    time.sleep(1.5)

    # 5. Cek dan tangani switch anomali di blok CATATAN
    # Aturan:
    # - Cari switch khusus "Tampilkan Anomali Usaha dan Keluarga" berdasarkan labelnya.
    # - Jika SUDAH dichecklist (checked/aria-checked): JANGAN dichecklist lagi.
    # - Jangan pernah mencentang "Anomali diselesaikan oleh admin".
    # - Jika BELUM dichecklist: baru klik switch "Tampilkan Anomali Usaha dan Keluarga".
    # - Jika kolom penjelasan masih kosong, isi dengan anomaly_explanation.
    catatan_res = page.evaluate(f"""(() => {{
        const rows = Array.from(document.querySelectorAll("div.tw\\\\:flex.tw\\\\:gap-2.tw\\\\:rounded-lg, div[class*='rounded-lg']"));

        // 1. Cari baris khusus "Tampilkan Anomali Usaha dan Keluarga" (hindari baris admin)
        const anomaliRow = rows.find(r => {{
            const t = (r.innerText || "").toLowerCase();
            return t.includes("tampilkan anomali") && !t.includes("admin");
        }});

        let switchAction = "NOT_FOUND";
        if (anomaliRow) {{
            const input = anomaliRow.querySelector("input[role='switch'], input[type='checkbox']");
            if (input) {{
                const isChecked = input.checked || 
                                  input.getAttribute("aria-checked") === "true" || 
                                  input.hasAttribute("data-checked");

                if (isChecked) {{
                    switchAction = "ALREADY_CHECKED";
                }} else {{
                    const control = anomaliRow.querySelector("[id*='control']") || input;
                    control.click();
                    switchAction = "CHECKED_NOW";
                }}
            }}
        }}

        // 2. Pastikan switch "Anomali diselesaikan oleh admin" TIDAK DICHECKLIST (jika sempat aktif, uncheck)
        const adminRow = rows.find(r => {{
            const t = (r.innerText || "").toLowerCase();
            return t.includes("anomali diselesaikan oleh admin");
        }});
        if (adminRow) {{
            const adminInput = adminRow.querySelector("input[role='switch'], input[type='checkbox']");
            if (adminInput) {{
                const isAdminChecked = adminInput.checked || 
                                       adminInput.getAttribute("aria-checked") === "true" || 
                                       adminInput.hasAttribute("data-checked");
                if (isAdminChecked) {{
                    const adminControl = adminRow.querySelector("[id*='control']") || adminInput;
                    adminControl.click();
                }}
            }}
        }}

        return {{ switchAction }};
    }})()""")

    time.sleep(1.0)

    # 6. Klik tombol Kirim utama
    page.evaluate("""(() => {
        const btn = Array.from(document.querySelectorAll("button")).find(b => b.innerText.trim() === "Kirim" && b.offsetWidth > 0);
        if (btn) btn.click();
    })()""")

    try:
        page.wait_for_selector("[class*='max-w-lg'], [role='dialog']", timeout=10000)
    except Exception:
        return False, "Modal ringkasan validasi tidak muncul"

    time.sleep(1.0)

    # 7. Cek validasi modal (harus 0 GALAT)
    val_res = page.evaluate("""(() => {
        const modal = document.querySelector("[class*='max-w-lg'], [role='dialog']");
        if (!modal) return { status: "ERR", msg: "Modal validasi tidak ditemukan" };
        const modalText = modal.innerText;
        const galatMatch = modalText.match(/GALAT\\s*\\n\\s*(\\d+)/i);
        const galatCount = galatMatch ? parseInt(galatMatch[1]) : 0;
        if (galatCount > 0) {
            const closeBtn = modal.querySelector("button[aria-label='Dismiss']") || 
                             Array.from(modal.querySelectorAll("button")).find(b => b.innerText.trim() === "Batal");
            if (closeBtn) closeBtn.click();
            return { status: "GALAT", count: galatCount, msg: galatCount + " galat ditemukan!" };
        }
        const btns = Array.from(modal.querySelectorAll("button")).filter(b => b.innerText.trim() === "Kirim");
        if (btns.length === 0) return { status: "ERR", msg: "Tombol Kirim di modal validasi tidak ditemukan" };
        btns[0].click();
        return { status: "OK" };
    })()""")

    if val_res.get("status") != "OK":
        return False, f"Validasi gagal: {val_res.get('msg')}"

    # 8. Klik tombol Konfirmasi pada dialog konfirmasi akhir
    try:
        page.wait_for_selector("button:has-text('Konfirmasi')", timeout=10000)
        page.locator("button:has-text('Konfirmasi')").first.click()
    except Exception:
        return False, "Tombol 'Konfirmasi' dialog akhir tidak ditemukan"

    # 9. Verifikasi hasil submit
    try:
        page.wait_for_function("""() => {
            const t = document.body ? document.body.innerText : "";
            return t.includes("Assignment Berhasil Disubmit") || t.includes("Perubahan Anda berhasil disimpan");
        }""", timeout=15000)
        return True, "Assignment Berhasil Disubmit"
    except Exception:
        return False, "Status sukses submit tidak terdeteksi pada layar"


def main():
    parser = argparse.ArgumentParser(description="Batch Automation Cleaning FASIH-SM SE2026 (1-Command)")
    parser.add_argument("--source", default="source/cleaning1.xlsx", help="File Excel sumber data")
    parser.add_argument("--start", type=int, default=1, help="Nomor baris awal untuk mulai diproses (default: 1)")
    parser.add_argument("--end", type=int, default=None, help="Nomor baris akhir untuk diproses")
    parser.add_argument("--limit", type=int, default=None, help="Jumlah baris maksimum yang diproses")
    parser.add_argument("--delay", type=float, default=1.5, help="Jeda waktu antar kuesioner dalam detik (default: 1.5)")
    parser.add_argument("--force", action="store_true", help="Paksa proses ulang semua baris meskipun sudah SUCCESS sebelumnya")
    parser.add_argument("--headless", action="store_true", help="Jalankan di latar belakang tanpa membuka jendela visual Chrome")
    parser.add_argument("--dry-run", action="store_true", help="Hanya tampilkan daftar link tanpa mengeksekusi")
    args = parser.parse_args()

    source_path = Path(args.source)
    all_tasks = read_excel_source(source_path)
    total_in_file = len(all_tasks)

    progress_data = load_progress()

    completed_indices = {int(k) for k, v in progress_data.items() if v.get("status") == "SUCCESS"}
    failed_indices = {int(k) for k, v in progress_data.items() if v.get("status") == "FAILED"}

    filtered_tasks = []
    skipped_count = 0
    for t in all_tasks:
        idx = t["index"]
        if idx < args.start:
            continue
        if args.end and idx > args.end:
            continue
        if not args.force and idx in completed_indices:
            skipped_count += 1
            continue
        filtered_tasks.append(t)
        if args.limit and len(filtered_tasks) >= args.limit:
            break

    console.print(
        Panel.fit(
            f"[bold cyan]⚡ FASIH-SM Batch Auto-Cleaning Engine (Standalone)[/bold cyan]\n"
            f"File Sumber      : [yellow]{source_path}[/yellow] (Total {total_in_file} baris)\n"
            f"SSO User         : [bold green]{config.SSO_USERNAME}[/bold green] (Auto-Login: {config.AUTO_LOGIN_SSO})\n"
            f"Profil Browser   : [magenta]{config.BROWSER_DATA_DIR}[/magenta] (Terisolasi & Di-ignore git)\n"
            f"Sudah Sukses     : [bold green]{len(completed_indices)} baris[/bold green] (Otomatis dilewati: {skipped_count})\n"
            f"Pernah Gagal     : [bold red]{len(failed_indices)} baris[/bold red] (Akan dicoba ulang)\n"
            f"Target Proses Ini: [bold cyan]{len(filtered_tasks)} baris[/bold cyan] (Mulai baris {args.start})\n"
            f"Tampilan Visual  : [yellow]{'Headless (Latar Belakang)' if args.headless else 'Google Chrome Jendela Visual'}[/yellow] | Delay: [cyan]{args.delay}s[/cyan]",
            border_style="cyan",
        )
    )

    if not filtered_tasks:
        console.print(
            "[bold green]✔ Semua baris target sudah SUCCESS sebelumnya! Tidak ada yang perlu diproses ulang.[/bold green]\n"
            "[dim]Gunakan opsi --force jika ingin tetap memproses ulang yang sudah sukses.[/dim]"
        )
        return

    if args.dry_run:
        table = Table(title="Daftar Baris Target (Dry Run)")
        table.add_column("No", justify="right", style="cyan")
        table.add_column("Link Kuesioner", style="magenta")
        table.add_column("Komentar KK", style="green")
        table.add_column("Status Terakhir", style="yellow")

        for t in filtered_tasks[:20]:
            prev_status = progress_data.get(str(t["index"]), {}).get("status", "PENDING")
            table.add_row(str(t["index"]), t["link"], t["komentar"], prev_status)

        console.print(table)
        if len(filtered_tasks) > 20:
            console.print(f"[dim]... dan {len(filtered_tasks) - 20} baris lainnya.[/dim]")
        return

    # Inisialisasi Google Chrome otomatis dengan Playwright
    console.print(f"[cyan]🚀 Meluncurkan Google Chrome otomatis ({'visual' if not args.headless else 'headless'})...[/cyan]")
    config.BROWSER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(config.BROWSER_DATA_DIR),
            channel=config.BROWSER_CHANNEL,
            headless=args.headless,
            viewport=None,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        page = context.pages[0] if context.pages else context.new_page()

        # Cek & pastikan status login awal
        try:
            page.goto("https://fasih-sm.bps.go.id/app/", wait_until="domcontentloaded", timeout=25000)
            time.sleep(2.0)
            if is_page_logged_out(page):
                console.print("[yellow]⚠ Status awal terdeteksi logout! Melakukan login otomatis SSO BPS...[/yellow]")
                if not perform_sso_login(page):
                    console.print("[bold red]❌ Gagal login otomatis SSO. Proses dibatalkan.[/bold red]")
                    context.close()
                    return
        except Exception as e:
            console.print(f"[dim yellow]Peringatan cek awal: {e}[/dim yellow]")

        success_count = 0
        fail_count = 0

        try:
            for i, item in enumerate(filtered_tasks, start=1):
                idx = item["index"]
                link = item["link"]
                komentar = item["komentar"]
                curr_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                console.print(f"\n[bold blue]━━━ [{i}/{len(filtered_tasks)}] Memproses Baris #{idx} ━━━[/bold blue]")
                console.print(f"Link    : [dim]{link}[/dim]")
                console.print(f"Komentar: [cyan]{komentar}[/cyan]")

                is_ok, msg = process_single_assignment(page, link, komentar)

                status_str = "SUCCESS" if is_ok else "FAILED"
                progress_data[str(idx)] = {
                    "link": link,
                    "komentar": komentar,
                    "status": status_str,
                    "detail": msg,
                    "timestamp": curr_time_str,
                }
                save_progress(progress_data)

                if is_ok:
                    success_count += 1
                    console.print(f"[bold green]✔ #{idx} Berhasil: {msg}[/bold green]")
                else:
                    fail_count += 1
                    console.print(f"[bold red]❌ #{idx} Gagal: {msg}[/bold red]")

                if i % 10 == 0 or i == len(filtered_tasks):
                    export_result_excel(all_tasks, progress_data, RESULT_EXCEL_FILE)

                if i < len(filtered_tasks) and args.delay > 0:
                    time.sleep(args.delay)

        except KeyboardInterrupt:
            console.print("\n[bold yellow]⚠ Eksekusi dihentikan pengguna (Ctrl+C). Checkpoint tersimpan aman![/bold yellow]")
        finally:
            export_result_excel(all_tasks, progress_data, RESULT_EXCEL_FILE)
            context.close()
            console.print(
                Panel.fit(
                    f"[bold cyan]📊 Rekapitulasi Eksekusi[/bold cyan]\n"
                    f"Selesai Diproses : [bold]{success_count + fail_count}[/bold] baris\n"
                    f"Berhasil (Submit): [bold green]{success_count}[/bold green]\n"
                    f"Gagal            : [bold red]{fail_count}[/bold red]\n"
                    f"Log Checkpoint   : [yellow]{PROGRESS_FILE}[/yellow]\n"
                    f"Hasil Excel      : [green]{RESULT_EXCEL_FILE}[/green]",
                    border_style="cyan",
                )
            )


if __name__ == "__main__":
    main()
