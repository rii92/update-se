#!/usr/bin/env python3
"""
Login FASIH-SM BPS (SE2026)
Meluncurkan Google Chrome dengan profil terisolasi (./chrome_data) dan
melakukan login otomatis SSO BPS menggunakan kredensial dari config.py.

Setelah login berhasil, browser dibiarkan terbuka agar sesi tersimpan pada
profil (./chrome_data), sehingga skrip cleaning berikutnya dapat langsung
dijalankan tanpa login ulang.

Contoh:
    python login_fasih.py
    python login_fasih.py --headless
    python login_fasih.py --close-after
"""

import argparse
import time
from pathlib import Path

from playwright.sync_api import Page, sync_playwright
from rich.console import Console
from rich.panel import Panel

import config

console = Console()

FASIH_APP_URL = f"{config.FASIH_BASE_URL}/app/"


def is_page_logged_out(page: Page) -> bool:
    """Mengecek apakah browser sedang berada pada status logout atau SSO."""
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
    """Melakukan login otomatis SSO BPS."""
    console.print(f"[bold yellow]🔐 Login otomatis SSO BPS (User: {username})...[/bold yellow]")
    try:
        if "sso.bps.go.id" not in page.url:
            sso_btn = page.locator(
                "a:has-text('Lanjutkan dengan SSO'), a:has-text('Login SSO BPS'), a[href*='oauth2/authorization'], .login-button"
            )
            if sso_btn.count() > 0 and sso_btn.first.is_visible():
                sso_btn.first.click()
            else:
                page.goto(config.SSO_LOGIN_URL, wait_until="domcontentloaded")
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


def main():
    parser = argparse.ArgumentParser(description="Login FASIH-SM BPS SE2026 (Auto-Login SSO)")
    parser.add_argument("--headless", action="store_true", help="Jalankan di latar belakang tanpa jendela Chrome")
    parser.add_argument("--close-after", action="store_true", help="Tutup browser otomatis setelah login berhasil")
    args = parser.parse_args()

    console.print(
        Panel.fit(
            f"[bold cyan]🔑 FASIH-SM Auto-Login SSO BPS[/bold cyan]\n"
            f"URL FASIH        : [yellow]{config.FASIH_BASE_URL}[/yellow]\n"
            f"SSO User         : [bold green]{config.SSO_USERNAME}[/bold green] (Auto-Login: {config.AUTO_LOGIN_SSO})\n"
            f"Profil Browser   : [magenta]{config.BROWSER_DATA_DIR}[/magenta]\n"
            f"Tampilan Visual  : [yellow]{'Headless (Latar Belakang)' if args.headless else 'Google Chrome Jendela Visual'}[/yellow]",
            border_style="cyan",
        )
    )

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

        console.print(f"[cyan]🚀 Membuka FASIH-SM: {FASIH_APP_URL}[/cyan]")

        try:
            page.goto(FASIH_APP_URL, wait_until="domcontentloaded", timeout=config.NAVIGATION_TIMEOUT)
            time.sleep(2.0)

            if is_page_logged_out(page):
                if not config.AUTO_LOGIN_SSO:
                    console.print("[bold red]❌ AUTO_LOGIN_SSO nonaktif. Silakan login manual di jendela Chrome.[/bold red]")
                elif not perform_sso_login(page):
                    console.print("[bold red]❌ Gagal login otomatis SSO. Periksa kredensial pada config.py.[/bold red]")
                    if not args.headless:
                        input("Tekan Enter untuk menutup browser...")
                    context.close()
                    return
            else:
                console.print("[bold green]✔ Sesi masih aktif, sudah login SSO BPS.[/bold green]")

            console.print(
                Panel.fit(
                    "[bold green]✔ Login FASIH-SM berhasil![/bold green]\n"
                    f"Sesi tersimpan pada profil: [magenta]{config.BROWSER_DATA_DIR}[/magenta]\n"
                    "[dim]Skrip cleaning (batch_cleaning.py / cleaning_nik.py) dapat langsung dijalankan.[/dim]",
                    border_style="green",
                )
            )
        except Exception as e:
            console.print(f"[bold red]❌ Terjadi kesalahan: {e}[/bold red]")

        if args.close_after or args.headless:
            context.close()
            console.print("[dim]Browser ditutup.[/dim]")
        else:
            console.print("[cyan]ℹ Browser dibiarkan terbuka. Tekan Ctrl+C untuk menutup sesi ini.[/cyan]")
            try:
                while True:
                    time.sleep(1.0)
            except KeyboardInterrupt:
                console.print("\n[yellow]Menutup browser...[/yellow]")
                context.close()


if __name__ == "__main__":
    main()
