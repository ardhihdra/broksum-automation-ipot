"""
IPOT - Broker Summary by Stock Automation
==========================================
Automates downloading Broker Summary by Stock CSV for a range of dates.

Requirements:
    pip install pywinauto pyautogui keyboard

Usage:
    1. Edit DATE_START and DATE_END at the bottom of this file.
    2. Run: python ipot_automation.py
    3. Keep your hands off the mouse/keyboard while it runs.
       Press Ctrl+C in the terminal to abort.

Notes:
    - Adjust SLEEP_* constants if your machine is slower/faster.
    - The script skips weekends automatically (no trading data).
    - If a modal or dialog appears unexpectedly, the script will
      pause and wait — check the screen.
"""

import time
import logging
import pyautogui
import pyperclip
from datetime import date, datetime, timedelta
import sys

# ─────────────────────────────────────────────
# CONFIGURATION — edit these
# ─────────────────────────────────────────────

APP_TITLE      = "IPOT"              # Window title keyword (partial match)
STOCK_CODES     = [
    "ADRO", "ANTM", "AADI", "ACRO",
    "BUVA", "BRPT", "BIPI", "BRMS", "BSDE", "BBCA", "BDKR", "BMTR", "BNBR", "BUMI",
    "CBDK", "CDIA", "COCO", "COCO", "CUAN", "CTTH", "CUAN"
    "DATA", "DAAZ", "DEWI", 
    "EMTK", "ENRG", "EXCL", "ELSA", "EMAS", "ESSA"
    "GGRM", "GOTO",
    "IMPC", "INCO", "INTP", "ITMG", "INDY", "ICON", "INET"
    "JARR",
    "MINA", "MBMA", "MBSS", "MINA",
    "PANI", "PGAS", "PTBA", "PSKT", "PTRO",
    "RMKO", "RMKE"
    "SUPA", "SGER", "SMBR", "SSIA", "SINI", "SMRA", 
    "TKIM", "TBIG", "TINS", "TLKM",
    "UNTR",
    "VKTR",
    "WBSA",
]            # Stock to query
SAVE_FOLDER    = r"g:\Ardhi\Activity Summary\INCo"  # Folder where CSVs are saved (must exist)

# ── Edit dates here ──────────────────────────────
DATE_START = date(2026, 5, 7)   # newer date (downloads in reverse)
DATE_END   = date(2026, 5, 7)    # older date

CENTER_X = 960
CENTER_y = 450

# Hardcoded context menu item — set after first run to skip the slow desktop scan.
# Leave all as "" to auto-discover (slow path will log the values for you to fill in).
SAVE_CSV_MENU_WINDOW_TITLE = "Context"             # title of the Menu window itself (usually "" for context menus)
SAVE_CSV_MENU_ITEM         = "Save To CSV"  # title of the menu item to click
SAVE_CSV_MENU_CONTROL_TYPE = ""     # control_type of the menu item

# Timing delays (seconds) — increase if the app is slow to respond
SLEEP_SHORT    = 0.45
SLEEP_MEDIUM   = 0.65
SLEEP_LONG     = 1.5

# ─────────────────────────────────────────────
# LOGGING SETUP
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)-5s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("ipot_debug.log", mode="w", encoding="utf-8"),
    ],
)
log = logging.getLogger("ipot")

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

_desktop_cache = None

def _get_desktop():
    global _desktop_cache
    if _desktop_cache is None:
        import pywinauto
        _desktop_cache = pywinauto.Desktop(backend="uia")
    return _desktop_cache

def pause(secs=SLEEP_MEDIUM):
    log.debug(f"pause({secs}s) ...")
    time.sleep(secs)

def type_text(text):
    """Select all existing text in the focused field then type new text."""
    log.debug(f"type_text: typing '{text}'")
    pyautogui.hotkey("ctrl", "a")
    pause(SLEEP_SHORT)
    pyautogui.typewrite(str(text), interval=0.05)
    pause(SLEEP_SHORT)
    log.debug("type_text: done")

def type_date(d: date):
    """Type a date into a segmented date field: day → Right → month → Right → year."""
    log.debug(f"type_date: entering {d}")
    log.debug(f"type_date: entering month {str(d.month)}")
    pyautogui.typewrite(str(d.month), interval=0.05)
    pyautogui.press("right")
    pause(SLEEP_SHORT)
    log.debug(f"type_date: entering day {str(d.day)}")
    pyautogui.typewrite(str(d.day), interval=0.05)
    pyautogui.press("right")
    pause(SLEEP_SHORT)
    log.debug(f"type_date: entering year {str(d.year)}")
    pyautogui.typewrite(str(d.year), interval=0.05)
    pause(SLEEP_SHORT)
    log.debug("type_date: done")

def fmt_date(d: date) -> str:
    """Format date as M/D/YYYY (IPOT format)."""
    result = d.strftime("%-m/%-d/%Y") if sys.platform != "win32" else d.strftime("%#m/%#d/%Y")
    log.debug(f"fmt_date({d}) → '{result}'")
    return result

def fmt_filename(stock_code: str, d: date) -> str:
    """Filename: YYYY-MM-DD_HH-MM-SS (safe for Windows)."""
    result = stock_code + "_" + d.strftime("%Y-%m-%d") + datetime.now().strftime("_%H-%M-%S")
    log.debug(f"fmt_filename({stock_code}, {d}) → '{result}'")
    return result

def is_weekday(d: date) -> bool:
    return d.weekday() < 5  # Mon=0 … Fri=4

def daterange(start: date, end: date):
    """Yield dates from start to end inclusive, weekdays only."""
    current = start
    while current >= end:
        if is_weekday(current):
            yield current
        current -= timedelta(days=1)

# ─────────────────────────────────────────────
# STEP FUNCTIONS
# ─────────────────────────────────────────────

def find_window():
    """Find and focus the IPOT window. Returns (app, win)."""
    import re
    import win32gui
    import pywinauto
    log.debug(f"find_window: scanning desktop for windows matching '.*{APP_TITLE}.*'")
    try:
        pattern = re.compile(f".*{APP_TITLE}.*", re.IGNORECASE)
        handles = []

        def _enum_cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if pattern.match(title):
                    handles.append(hwnd)

        win32gui.EnumWindows(_enum_cb, None)

        log.debug(f"find_window: total matches found = {len(handles)}")
        for i, h in enumerate(handles):
            log.debug(f"  [{i}] title='{win32gui.GetWindowText(h)}'  handle={h}")

        if not handles:
            raise Exception("No IPOT windows found")

        # Wrap handles as lightweight pywinauto window specs for compatibility
        desktop = _get_desktop()
        matches = [desktop.window(handle=h) for h in handles]

        # Prefer a visible, enabled window; fall back to first match
        target = next(
            (w for w in matches if w.is_visible() and w.is_enabled()),
            matches[0]
        )
        log.info(f"find_window: selected window '{target.window_text()}' (handle={target.handle})")

        app = pywinauto.Application(backend="uia").connect(handle=target.handle)
        win = app.window(handle=target.handle)
        log.debug("find_window: connected via handle, calling set_focus()")
        win.set_focus()
        pause(SLEEP_MEDIUM)
        log.info(f"find_window: focused on '{win.window_text()}'")
        return app, win

    except Exception as e:
        log.error(f"find_window: failed — {e}")
        log.info("Make sure IPOT is running, then press Enter to retry...")
        input()
        return find_window()


def open_broker_summary(win):
    """Click Security Analysis → Broker Summary by Stock."""
    log.info("open_broker_summary: attempting menu 'Security Analysis->Broker Summary by Stock'")
    try:
        win.menu_select("Security Analysis->Broker Summary by Stock")
        pause(SLEEP_LONG)
        log.info("open_broker_summary: opened via menu_select")
    except Exception as e:
        log.warning(f"open_broker_summary: menu_select failed ({e}), falling back to keyboard")
        win.set_focus()
        pause(SLEEP_SHORT)
        log.debug("open_broker_summary: pressing Alt to activate menu bar")
        pyautogui.hotkey("alt")
        pause(SLEEP_SHORT)
        log.debug("open_broker_summary: pressing Right to reach Security Analysis")
        pyautogui.press("right")
        pause(SLEEP_SHORT)
        log.debug("open_broker_summary: pressing Enter to open Security Analysis")
        pyautogui.press("enter")
        pause(SLEEP_SHORT)
        log.debug("open_broker_summary: pressing Down to reach Broker Summary by Stock")
        pyautogui.press("down")
        pause(SLEEP_SHORT)
        log.debug("open_broker_summary: pressing Enter to select Broker Summary by Stock")
        pyautogui.press("enter")
        pause(SLEEP_LONG)
        log.info("open_broker_summary: opened via keyboard fallback")


def set_fields(win, stock: str, query_date: date):
    """Open Broker Summary dialog and fill fields via keyboard — dialog opens focused on stock field."""
    date_str = fmt_date(query_date)
    log.info(f"set_fields: stock='{stock}'  date='{date_str}'")
    open_broker_summary(win)
    type_text(stock)
    pyautogui.press("tab")
    type_date(query_date)
    pyautogui.press("tab")
    type_date(query_date)
    pyautogui.press("tab")
    pause(SLEEP_SHORT)
    pyautogui.press("n")
    log.info(f"set_fields: done — {stock} | {date_str}")


def trigger_search(win):
    """Press Enter to submit, then wait for the DataGrid to be ready instead of sleeping."""
    log.info("trigger_search: pressing Enter to submit")
    pyautogui.press("enter")
    log.debug("trigger_search: waiting for DataGrid (max 10s)")
    # try:
    #     win.child_window(control_type="DataGrid").wait("ready", timeout=10)
    #     log.info("trigger_search: DataGrid ready")
    # except Exception as e:
    #     log.warning(f"trigger_search: DataGrid wait failed ({e}), falling back to fixed sleep")
    #     pause(SLEEP_LONG)
    pause(SLEEP_LONG)
    log.info("trigger_search: done waiting for results to load")


def save_to_csv(win, stock_code: str, query_date: date):
    """Right-click the data grid and choose Save to CSV."""
    filename = fmt_filename(stock_code, query_date)
    log.info(f"save_to_csv: target filename='{filename}'")

    log.debug("save_to_csv: looking for DataGrid control")
    try:
        grid = win.child_window(control_type="DataGrid")
        if not CENTER_X or not CENTER_y:
            rect = grid.rectangle()
            height = rect.bottom - rect.top
            cx = (rect.left + rect.right) // 2
            cy = (rect.top + rect.bottom)
        else:
            cx = CENTER_X
            cy = CENTER_y
        log.debug(f"save_to_csv: DataGrid rect=top:{rect.top} bottom:{rect.bottom} left:{rect.left} right:{rect.right} height:{height}")
        log.debug(f"save_to_csv: moving mouse to ({cx},{cy})")
        pyautogui.moveTo(cx, cy)
        pause(SLEEP_SHORT)
        log.debug("save_to_csv: right-clicking DataGrid")
        grid.right_click_input()
        log.info("save_to_csv: right-clicked DataGrid")
    except Exception as e:
        rect = win.rectangle()
        height = rect.bottom - rect.top
        cx = (rect.left + rect.right) // 2
        # cy = rect.top + (rect.bottom - rect.top) // 2
        cy = 450
        log.warning(f"save_to_csv: DataGrid not found ({e})")
        log.warning(f"save_to_csv: window rect=top:{rect.top} bottom:{rect.bottom} left:{rect.left} right:{rect.right} height:{height}")
        log.warning(f"save_to_csv: moving mouse to ({cx},{cy})")
        pyautogui.moveTo(cx, cy)
        pause(SLEEP_SHORT)
        pyautogui.rightClick(cx, cy)

    pause(SLEEP_SHORT)  # let context menu render

    # Primary: move mouse right into the menu, then down to the 3rd item and click.
    click_x, click_y = pyautogui.position()
    menu_x = click_x + 15            # step right into the menu body
    menu_y = click_y + 22 * 1 + 11   # centre of 3rd item (~22 px/item, 0-based index 2)
    log.debug(f"save_to_csv: mouse nav — post-click=({click_x},{click_y}), 3rd item≈({menu_x},{menu_y})")
    try:
        pyautogui.moveTo(menu_x, menu_y)
        pause(SLEEP_SHORT)
        pyautogui.click()
        log.info("save_to_csv: clicked 3rd menu item via mouse navigation")
    except Exception as e:
        log.warning(f"save_to_csv: mouse nav failed ({e}), falling back to UIA scan")
        desktop = _get_desktop()
        try:
            if SAVE_CSV_MENU_ITEM:
                log.debug(
                    f"save_to_csv: UIA fast path — "
                    f"title='{SAVE_CSV_MENU_WINDOW_TITLE}'  item='{SAVE_CSV_MENU_ITEM}'"
                )
                ctx_menu = desktop.window(title=SAVE_CSV_MENU_WINDOW_TITLE, control_type="Menu")
                target_item = ctx_menu.child_window(title_re=f".*{SAVE_CSV_MENU_ITEM}.*")
            else:
                log.debug("save_to_csv: UIA slow path — scanning desktop for Menu window")
                ctx_menu = desktop.window(title_re=".*", control_type="Menu")
                log.info(
                    f"save_to_csv: Menu found — title='{ctx_menu.window_text()}'  "
                    f"(set SAVE_CSV_MENU_WINDOW_TITLE to this value)"
                )
                try:
                    for item in ctx_menu.children():
                        log.debug(
                            f"  title='{item.window_text()}'"
                            f"  control_type='{item.element_info.control_type}'"
                        )
                except Exception:
                    pass
                target_item = ctx_menu.child_window(title_re=".*Save.*CSV.*")
            log.info(f"save_to_csv: UIA found '{target_item.window_text()}', clicking")
            target_item.click_input()
            log.info("save_to_csv: clicked via UIA fallback")
        except Exception as e2:
            log.warning(f"save_to_csv: UIA fallback also failed ({e2}), pressing 'S'")
            pyautogui.hotkey("s")

    pause(SLEEP_MEDIUM)
    _handle_save_dialog(filename)


def _handle_save_dialog(filename: str):
    """Handle the Save file dialog — filename field is already focused on open."""
    log.info(f"_handle_save_dialog: filename='{filename}'")
    pause(SLEEP_SHORT)  # let focus settle on the filename field
    pyautogui.hotkey("ctrl", "a")
    pause(SLEEP_SHORT)
    pyautogui.typewrite(filename, interval=0.05)
    pyautogui.press("enter")

    # """Handle the Save file dialog: set filename and click Save."""
    # log.info(f"_handle_save_dialog: saved as '{filename}.csv'")
    # import pywinauto
    # log.info(f"_handle_save_dialog: waiting for Save dialog, filename='{filename}'")
    # pause(SLEEP_MEDIUM)

    # try:
    #     save_dlg = pywinauto.Desktop(backend="uia").window(
    #         title_re="(Save|Save As|另存为|저장)"
    #     )
    #     log.debug(f"_handle_save_dialog: dialog found — title='{save_dlg.window_text()}'")
    #     save_dlg.set_focus()
    #     pause(SLEEP_SHORT)

    #     log.debug("_handle_save_dialog: locating filename Edit field")
    #     fname_field = save_dlg.child_window(
    #         control_type="Edit", title_re="(File name|FileName|文件名)"
    #     )
    #     log.debug(f"_handle_save_dialog: current filename field value='{fname_field.window_text()}'")
    #     fname_field.set_focus()
    #     fname_field.set_edit_text(filename)
    #     log.debug(f"_handle_save_dialog: filename set to '{filename}'")
    #     pause(SLEEP_SHORT)

    #     log.debug("_handle_save_dialog: clicking Save/OK button")
    #     save_dlg.child_window(
    #         title_re="(Save|OK|确定)", control_type="Button"
    #     ).click_input()
    #     pause(SLEEP_MEDIUM)
    #     log.info(f"_handle_save_dialog: saved as '{filename}.csv'")

    # except Exception as e:
    #     log.warning(f"_handle_save_dialog: dialog not found via pywinauto ({e}), using keyboard fallback")
    #     pyautogui.hotkey("ctrl", "a")
    #     pause(SLEEP_SHORT)
    #     type_text(filename)
    #     pyautogui.press("enter")
    #     pause(SLEEP_MEDIUM)
    #     log.info(f"_handle_save_dialog: saved as '{filename}.csv' (keyboard fallback)")


# ─────────────────────────────────────────────
# MAIN LOOP
# ─────────────────────────────────────────────

def run(stock_code: str,start_date: date, end_date: date):
    """
    Download Broker Summary CSVs for every weekday from start_date down to end_date.

    start_date : more recent date  (e.g. date(2026, 4, 16))
    end_date   : older date        (e.g. date(2026, 4, 1))
    """
    log.info("=" * 50)
    log.info("IPOT Broker Summary Automation")
    log.info(f"Range: {start_date} → {end_date}")
    log.info(f"Stock: {stock_code}")
    log.info(f"Save folder: {SAVE_FOLDER}")
    log.info("=" * 50)

    # Fail-safe: move mouse to top-left corner to abort
    pyautogui.FAILSAFE = True
    log.debug("pyautogui FAILSAFE enabled (move mouse to top-left corner to abort)")

    log.info("Step 1: finding IPOT window")
    app, win = find_window()
    # log.info("  Step 2: opening Broker Summary by Stock panel")
    # open_broker_summary(win)

    dates = list(daterange(start_date, end_date))
    total = len(dates)
    log.info(f"Dates to process: {total}  ({dates[0] if dates else '—'} → {dates[-1] if dates else '—'})")

    for i, d in enumerate(dates, 1):
        log.info(f"{'─'*40}")
        log.info(f"[{i}/{total}] Processing {d}")
        t0 = time.time()
        try:

            log.info("  Step A: set_fields")
            set_fields(win, stock_code, d)

            log.info("  Step B: trigger_search")
            trigger_search(win)

            log.info("  Step C: save_to_csv")
            save_to_csv(win, stock_code, d)

            elapsed = time.time() - t0
            log.info(f"  Done in {elapsed:.1f}s")
            pause(SLEEP_MEDIUM)

        except KeyboardInterrupt:
            log.warning("Aborted by user (KeyboardInterrupt)")
            sys.exit(0)
        except Exception as e:
            log.error(f"Error on {d}: {e}", exc_info=True)
            log.warning("Skipping this date — check the screen")
            pause(SLEEP_LONG)
        finally:
            log.info("  Closing Broker Summary panel")
            win.set_focus()
            pause(SLEEP_SHORT)
            pyautogui.hotkey("alt", "f4")
            pause(SLEEP_SHORT)
            pyautogui.moveTo(CENTER_X, 10)
            pyautogui.click()

    log.info("=" * 50)
    log.info("All done!")


if __name__ == "__main__":
    # ─────────────────────────────────────────────────
    log.info("Starting. Please open IPOT and click the header!")
    pause(SLEEP_LONG)
    for stock_code in STOCK_CODES:
        run(stock_code, DATE_START, DATE_END)
