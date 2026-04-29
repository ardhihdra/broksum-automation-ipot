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
from datetime import date, timedelta
import sys

# ─────────────────────────────────────────────
# CONFIGURATION — edit these
# ─────────────────────────────────────────────

APP_TITLE      = "IPOT"              # Window title keyword (partial match)
STOCK_CODE     = "MBMA"             # Stock to query
SAVE_FOLDER    = r"g:\Ardhi\Activity Summary\INCo"  # Folder where CSVs are saved (must exist)

# Hardcoded context menu item — set after first run to skip the slow desktop scan.
# Leave all as "" to auto-discover (slow path will log the values for you to fill in).
SAVE_CSV_MENU_WINDOW_TITLE = "Context"             # title of the Menu window itself (usually "" for context menus)
SAVE_CSV_MENU_ITEM         = "Save To CSV"  # title of the menu item to click
SAVE_CSV_MENU_CONTROL_TYPE = ""     # control_type of the menu item

# Timing delays (seconds) — increase if the app is slow to respond
SLEEP_SHORT    = 0.5
SLEEP_MEDIUM   = 0.8
SLEEP_LONG     = 2.5

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

def fmt_filename(d: date) -> str:
    """Filename: YYYY-MM-DD (safe for Windows)."""
    result = d.strftime("%Y-%m-%d")
    log.debug(f"fmt_filename({d}) → '{result}'")
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
        desktop = pywinauto.Desktop(backend="uia")
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
    """Fill in the three input fields: stock code, from-date, to-date."""
    date_str = fmt_date(query_date)
    log.info(f"set_fields: stock='{stock}'  date='{date_str}'")

    win.set_focus()
    pause(SLEEP_SHORT)

    log.debug("set_fields: enumerating Edit controls in window")
    try:
        fields = win.children(control_type="Edit")
        log.debug(f"set_fields: found {len(fields)} Edit field(s)")
        for i, f in enumerate(fields):
            log.debug(
                f"  field[{i}] title='{f.window_text()}'"
                f"  auto_id='{f.automation_id()}'"
                f"  value='{f.get_value() if hasattr(f, 'get_value') else '?'}'"
            )

        if len(fields) >= 3:
            log.debug(f"set_fields: setting field[0] (stock) → '{stock}'")
            fields[0].set_focus()
            fields[0].set_edit_text(stock)
            pause(SLEEP_SHORT)

            log.debug(f"set_fields: setting field[1] (from-date) → '{date_str}'")
            fields[1].set_focus()
            fields[1].set_edit_text(date_str)
            pause(SLEEP_SHORT)

            log.debug(f"set_fields: setting field[2] (to-date) → '{date_str}'")
            fields[2].set_focus()
            fields[2].set_edit_text(date_str)
            pause(SLEEP_SHORT)

            # Tab moves focus to Value/Net select — press N to choose Net
            log.debug("set_fields: Tab → Value/Net select, pressing N to choose Net")
            pyautogui.press("tab")
            pause(SLEEP_SHORT)
            pyautogui.press("n")

            log.info(f"set_fields: all fields set — {stock} | {date_str} | {date_str}")
            return
        else:
            log.warning(f"set_fields: only {len(fields)} Edit field(s) found, need ≥3")
            raise Exception(f"expected ≥3 Edit fields, got {len(fields)}")

    except Exception as e:
        log.warning(f"set_fields: direct field access failed ({e}), switching to Tab fallback")

    # Fallback: Tab through fields
    # log.debug("set_fields: pressing Alt+F4 to close any stray dialog")
    # pyautogui.hotkey("alt", "F4")
    # pause(SLEEP_SHORT)
    log.debug("set_fields: re-opening Broker Summary dialog")
    open_broker_summary(win)
    log.debug("set_fields: Tab → field 1, typing stock code")
    # pyautogui.press("tab")
    type_text(stock)
    log.debug("set_fields: Tab → field 2, typing from-date")
    pyautogui.press("tab")
    type_date(query_date)
    log.debug("set_fields: Tab → field 3, typing to-date")
    pyautogui.press("tab")
    type_date(query_date)

    # Tab moves focus to Value/Net select — press N to choose Net
    log.debug("set_fields: Tab → Value/Net select, pressing N to choose Net (fallback path)")
    pyautogui.press("tab")
    pause(SLEEP_SHORT)
    pyautogui.press("n")
    log.info(f"set_fields: fields set via Tab fallback — {stock} | {date_str} | {date_str}")


def trigger_search(win):
    """Press Enter or click Search/OK to load the data."""
    log.info("trigger_search: looking for Search/OK/Go/Load button")
    pyautogui.press("enter")
    # try:
    #     btn = win.child_window(title_re="(Search|OK|Go|Load)", control_type="Button")
    #     log.debug(f"trigger_search: found button '{btn.window_text()}', clicking")
    #     btn.click_input()
    #     log.info("trigger_search: button clicked")
    # except Exception as e:
    #     log.warning(f"trigger_search: button not found ({e}), pressing Enter as fallback")
    #     pyautogui.press("enter")
    #     log.info("trigger_search: Enter pressed")
    pause(SLEEP_LONG)
    log.info("trigger_search: done waiting for results to load")


def save_to_csv(win, query_date: date):
    """Right-click the data grid and choose Save to CSV."""
    filename = fmt_filename(query_date)
    log.info(f"save_to_csv: target filename='{filename}'")

    log.debug("save_to_csv: looking for DataGrid control")
    try:
        grid = win.child_window(control_type="DataGrid")
        rect = grid.rectangle()
        height = rect.bottom - rect.top
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom)
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

    pause(SLEEP_MEDIUM)

    log.debug("save_to_csv: looking for context menu with 'Save.*CSV' item")
    try:
        import pywinauto
        if SAVE_CSV_MENU_ITEM:
            # Fast path: find the Menu window by exact title (context menus usually have
            # an empty title so this is much faster than title_re=".*"), then find the
            # known item inside it as a child — avoids the full desktop UIA scan.
            log.debug(
                f"save_to_csv: fast path — "
                f"menu_window_title='{SAVE_CSV_MENU_WINDOW_TITLE}'  "
                f"item='{SAVE_CSV_MENU_ITEM}'  control_type='{SAVE_CSV_MENU_CONTROL_TYPE}'"
            )
            ctx_menu = pywinauto.Desktop(backend="uia").window(
                title=SAVE_CSV_MENU_WINDOW_TITLE, control_type="Menu"
            )
            log.debug(
                    f"save_to_csv: context menu found — title='{ctx_menu.window_text()}'"   
                    f"  control_type='{ctx_menu.element_info.control_type}'"
                    f"  auto_id='{ctx_menu.automation_id()}'"
                    f"  class='{ctx_menu.friendly_class_name()}'"
            )
            # Use title_re for child lookup — the item's stored text may include
            # an ampersand accelerator (e.g. "Save to &CSV") that breaks exact match.
            target_item = ctx_menu.child_window(
                title_re=f".*{SAVE_CSV_MENU_ITEM}.*"
            )
        else:
            # Slow path: scan desktop for any Menu window, then find item by regex.
            # Run this once to discover the values, then hardcode them above.
            log.debug("save_to_csv: slow path — scanning desktop for Menu window")
            ctx_menu = pywinauto.Desktop(backend="uia").window(
                title_re=".*", control_type="Menu"
            )
            log.info(
                f"save_to_csv: Menu window found — "
                f"title='{ctx_menu.window_text()}'  "
                f"(set SAVE_CSV_MENU_WINDOW_TITLE to this value)"
            )
            # Log all menu items for debugging
            try:
                items = ctx_menu.children()
                log.debug(f"save_to_csv: context menu has {len(items)} item(s):")
                for item in items:
                    log.debug(
                        f"  title='{item.window_text()}'"
                        f"  control_type='{item.element_info.control_type}'"
                        f"  auto_id='{item.automation_id()}'"
                        f"  class='{item.friendly_class_name()}'"
                    )
            except Exception:
                pass
            target_item = ctx_menu.child_window(title_re=".*Save.*CSV.*")

        log.info(
            f"save_to_csv: found menu item"
            f"  title='{target_item.window_text()}'"
            f"  control_type='{target_item.element_info.control_type}'"
            f"  auto_id='{target_item.automation_id()}'"
            f"  class='{target_item.friendly_class_name()}'"
        )
        log.debug(f"save_to_csv: clicking menu item '{target_item.window_text()}'")
        target_item.click_input()
        log.info("save_to_csv: 'Save to CSV' menu item clicked")
    except Exception as e:
        log.warning(f"save_to_csv: context menu approach failed ({e}), pressing 'S' as shortcut")
        pyautogui.hotkey("s")

    pause(SLEEP_MEDIUM)
    _handle_save_dialog(filename)


def _handle_save_dialog(filename: str):
    """Handle the Save file dialog: set filename and click Save."""
    import pywinauto
    log.info(f"_handle_save_dialog: waiting for Save dialog, filename='{filename}'")
    pause(SLEEP_MEDIUM)

    try:
        save_dlg = pywinauto.Desktop(backend="uia").window(
            title_re="(Save|Save As|另存为|저장)"
        )
        log.debug(f"_handle_save_dialog: dialog found — title='{save_dlg.window_text()}'")
        save_dlg.set_focus()
        pause(SLEEP_SHORT)

        log.debug("_handle_save_dialog: locating filename Edit field")
        fname_field = save_dlg.child_window(
            control_type="Edit", title_re="(File name|FileName|文件名)"
        )
        log.debug(f"_handle_save_dialog: current filename field value='{fname_field.window_text()}'")
        fname_field.set_focus()
        fname_field.set_edit_text(filename)
        log.debug(f"_handle_save_dialog: filename set to '{filename}'")
        pause(SLEEP_SHORT)

        log.debug("_handle_save_dialog: clicking Save/OK button")
        save_dlg.child_window(
            title_re="(Save|OK|确定)", control_type="Button"
        ).click_input()
        pause(SLEEP_MEDIUM)
        log.info(f"_handle_save_dialog: saved as '{filename}.csv'")

    except Exception as e:
        log.warning(f"_handle_save_dialog: dialog not found via pywinauto ({e}), using keyboard fallback")
        pyautogui.hotkey("ctrl", "a")
        pause(SLEEP_SHORT)
        type_text(filename)
        pyautogui.press("enter")
        pause(SLEEP_MEDIUM)
        log.info(f"_handle_save_dialog: saved as '{filename}.csv' (keyboard fallback)")


# ─────────────────────────────────────────────
# MAIN LOOP
# ─────────────────────────────────────────────

def run(start_date: date, end_date: date):
    """
    Download Broker Summary CSVs for every weekday from start_date down to end_date.

    start_date : more recent date  (e.g. date(2026, 4, 16))
    end_date   : older date        (e.g. date(2026, 4, 1))
    """
    log.info("=" * 50)
    log.info("IPOT Broker Summary Automation")
    log.info(f"Range: {start_date} → {end_date}")
    log.info(f"Stock: {STOCK_CODE}")
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
            set_fields(win, STOCK_CODE, d)

            log.info("  Step B: trigger_search")
            trigger_search(win)

            log.info("  Step C: save_to_csv")
            save_to_csv(win, d)

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

    log.info("=" * 50)
    log.info("All done!")


if __name__ == "__main__":
    # ── Edit dates here ──────────────────────────────
    DATE_START = date(2026, 4, 19)   # newer date (downloads in reverse)
    DATE_END   = date(2026, 4, 1)    # older date
    # ─────────────────────────────────────────────────
    log.info("Starting. Please open IPOT and click the header!")
    pause(SLEEP_LONG)
    run(DATE_START, DATE_END)
