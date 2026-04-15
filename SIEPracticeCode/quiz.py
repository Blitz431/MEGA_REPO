"""
SIE Exam Quizzer
----------------
Reads questions from "QUESTION BANK AND TRACKER.xlsx" and quizzes the user
via a CLI. Writes RIGHT / WRONG / # times asked back to the Excel file.

Usage:  python quiz.py
"""

import os
import sys
import csv
import random
import datetime

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl is not installed. Run:  pip install openpyxl")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "QUESTION BANK AND TRACKER.xlsx")
LOG_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "session_log.csv")

SECTION_MAP = {
    "1": {
        "name": "Knowledge of Capital Markets",
        "subsections": ["1.1", "1.2", "1.3"],
        "sub_names": {
            "1.1": "Regulatory Entities, Agencies & Market Participants",
            "1.2": "Market Structure",
            "1.3": "Economic Factors",
        },
    },
    "2": {
        "name": "Understanding Products and Their Risks",
        "subsections": ["2.1", "2.2", "2.3", "2.4", "2.5"],
        "sub_names": {
            "2.1": "Equity Securities",
            "2.2": "Debt Securities",
            "2.3": "Packaged Products",
            "2.4": "Options",
            "2.5": "Alternative Investments & Other Products",
        },
    },
    "3": {
        "name": "Understanding Trading, Customer Accounts & Prohibited Activities",
        "subsections": ["3.1", "3.2", "3.3"],
        "sub_names": {
            "3.1": "Trading Securities",
            "3.2": "Customer Accounts & Compliance",
            "3.3": "Prohibited Activities",
        },
    },
    "4": {
        "name": "Overview of the Regulatory Framework",
        "subsections": ["4.1", "4.2", "4.3"],
        "sub_names": {
            "4.1": "SROs and Regulatory Bodies",
            "4.2": "Registration & Licensing",
            "4.3": "Key Regulations & Rules",
        },
    },
}

# Real SIE exam distribution
EXAM_DISTRIBUTION = {"1": 12, "2": 33, "3": 23, "4": 7}

PASS_THRESHOLD = 0.72  # 72% to pass the SIE

DIVIDER  = "=" * 62
THIN_DIV = "-" * 62

# ---------------------------------------------------------------------------
# Help text
# ---------------------------------------------------------------------------

HELP_TEXT = """
╔══════════════════════════════════════════════════════════════╗
║                    SIE QUIZZER — HELP                       ║
╠══════════════════════════════════════════════════════════════╣
║  MODES                                                      ║
║  F  Full Exam    75 questions, real SIE distribution        ║
║  S  Study Mode   pick sections/subsections + target count   ║
║  W  Weak Areas   only questions you've answered wrong       ║
╠══════════════════════════════════════════════════════════════╣
║  CHECKLIST MENUS                                            ║
║  <number>  toggle that item on/off                          ║
║  A         select all                                       ║
║  C         clear all                                        ║
║  Enter     confirm selection                                ║
╠══════════════════════════════════════════════════════════════╣
║  ANSWERING QUESTIONS                                        ║
║  1–4       pick your answer (choices are shuffled each time)║
║  H or ?    show this help screen                            ║
║  Q         quit to main menu                                ║
║  Ctrl+C    exit the program                                 ║
╠══════════════════════════════════════════════════════════════╣
║  TRACKING                                                   ║
║  RIGHT / WRONG / # times asked are saved to Excel after     ║
║  each answer. Close Excel before running to allow saving.   ║
╚══════════════════════════════════════════════════════════════╝
"""

# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------

def load_all_questions() -> list:
    """Load every question from the Excel file. Returns list of dicts."""
    if not os.path.exists(EXCEL_PATH):
        print(f"\nERROR: Excel file not found at:\n  {EXCEL_PATH}", flush=True)
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except Exception as exc:
        print(f"\nERROR: Could not open Excel file: {exc}", flush=True)
        sys.exit(1)

    ws = wb.active
    questions = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) < 10:
            continue

        section_raw, question, choice_a, choice_b, choice_c, choice_d, \
            correct, times, right, wrong = (row[i] for i in range(10))

        if section_raw is None:
            continue

        # Normalize section to "X.X" string (it is stored as a float)
        try:
            section_str = f"{float(section_raw):.1f}"
        except (ValueError, TypeError):
            section_str = str(section_raw).strip()

        # Skip rows with missing data
        if any(v is None for v in [question, choice_a, choice_b, choice_c, choice_d, correct]):
            print(f"  [Warning] Skipping incomplete row {row_idx}", flush=True)
            continue

        correct = str(correct).strip().upper()
        if correct not in {"A", "B", "C", "D"}:
            print(f"  [Warning] Row {row_idx} has invalid correct answer '{correct}' — skipping", flush=True)
            continue

        questions.append({
            "row":        row_idx,
            "section":    section_str,
            "question":   str(question),
            "choices":    {
                "A": str(choice_a),
                "B": str(choice_b),
                "C": str(choice_c),
                "D": str(choice_d),
            },
            "correct":    correct,
            "times_asked": int(times)  if times  is not None else 0,
            "right":       int(right)  if right   is not None else 0,
            "wrong":       int(wrong)  if wrong   is not None else 0,
        })

    wb.close()
    return questions


def save_result(row: int, correct: bool) -> None:
    """Increment tracking columns in the Excel file for the given row."""
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active

        def inc(col):
            v = ws.cell(row=row, column=col).value
            ws.cell(row=row, column=col).value = (int(v) if v is not None else 0) + 1

        inc(8)  # # Of times asked
        if correct:
            inc(9)   # RIGHT
        else:
            inc(10)  # WRONG

        wb.save(EXCEL_PATH)
        wb.close()
    except PermissionError:
        print("\n  [Warning] Could not save to Excel — close the file in Excel and re-run to save stats.")
    except Exception as exc:
        print(f"\n  [Warning] Could not save result: {exc}")


def append_session_log(mode: str, sections: str, target: int,
                       correct: int, total: int) -> None:
    """Append one line to session_log.csv."""
    pct = (correct / total * 100) if total > 0 else 0.0
    write_header = not os.path.exists(LOG_PATH)
    try:
        with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(["date", "mode", "sections", "target",
                                 "correct", "total", "pct"])
            writer.writerow([
                datetime.date.today().isoformat(),
                mode, sections, target, correct, total,
                f"{pct:.1f}",
            ])
    except Exception:
        pass  # Don't crash if log write fails

# ---------------------------------------------------------------------------
# Adaptive weighting
# ---------------------------------------------------------------------------

def question_weight(q: dict) -> float:
    """Higher weight = shown more often (weak questions prioritized)."""
    wrong = q["wrong"]
    right = q["right"]
    total = wrong + right
    if total == 0:
        return 2.0  # unasked questions get medium-high priority
    accuracy = right / total
    return max(0.5, 2.0 - accuracy * 1.5)


def weighted_sample(pool: list, k: int) -> list:
    """Sample k items from pool using adaptive weights (with replacement de-duped)."""
    if k >= len(pool):
        result = pool[:]
        random.shuffle(result)
        return result
    weights = [question_weight(q) for q in pool]
    # Sample without replacement by drawing one at a time
    selected = []
    remaining = list(zip(pool, weights))
    for _ in range(k):
        if not remaining:
            break
        items, wts = zip(*remaining)
        chosen = random.choices(items, weights=wts, k=1)[0]
        selected.append(chosen)
        remaining = [(it, w) for it, w in remaining if it is not chosen]
    return selected

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

def show_dashboard(all_questions: list) -> None:
    """Print a progress table grouped by subsection."""
    # Collect all known subsections in canonical order
    ordered_subs = []
    for sk in sorted(SECTION_MAP.keys()):
        ordered_subs.extend(SECTION_MAP[sk]["subsections"])

    # Group questions by subsection
    by_sub = {s: [] for s in ordered_subs}
    for q in all_questions:
        if q["section"] in by_sub:
            by_sub[q["section"]].append(q)

    print(f"\n{DIVIDER}")
    print("  YOUR PROGRESS")
    print(DIVIDER)
    print(f"  {'Section':<8} {'Topic':<42} {'Q':>4} {'Right':>5} {'Wrong':>5} {'%':>6}")
    print(THIN_DIV)

    for sk in sorted(SECTION_MAP.keys()):
        for sub in SECTION_MAP[sk]["subsections"]:
            qs     = by_sub.get(sub, [])
            total  = len(qs)
            right  = sum(q["right"]  for q in qs)
            wrong  = sum(q["wrong"]  for q in qs)
            asked  = right + wrong
            pct    = f"{right/asked*100:.0f}%" if asked > 0 else "—"
            name   = SECTION_MAP[sk]["sub_names"].get(sub, "")
            print(f"  {sub:<8} {name:<42} {total:>4} {right:>5} {wrong:>5} {pct:>6}")

    print(THIN_DIV)
    total_q = len(all_questions)
    total_r = sum(q["right"] for q in all_questions)
    total_w = sum(q["wrong"] for q in all_questions)
    asked   = total_r + total_w
    pct_all = f"{total_r/asked*100:.0f}%" if asked > 0 else "—"
    print(f"  {'TOTAL':<8} {'':<42} {total_q:>4} {total_r:>5} {total_w:>5} {pct_all:>6}")
    print(DIVIDER)

# ---------------------------------------------------------------------------
# Menus
# ---------------------------------------------------------------------------

def _input(prompt: str) -> str:
    """Wrapper around input() that prints help on H/? and returns stripped upper."""
    while True:
        try:
            val = input(prompt).strip()
        except EOFError:
            return "Q"
        if val.upper() in ("H", "?"):
            print(HELP_TEXT)
        else:
            return val.upper()


def reset_data() -> None:
    """Zero out all tracking columns in Excel and delete the session log."""
    print(f"\n{DIVIDER}")
    print("  RESET DATA")
    print(DIVIDER)
    print("  This will clear ALL right/wrong/times-asked stats in the")
    print("  question bank and delete session_log.csv.")
    print()
    confirm = input('  Type  RESET DATA  to confirm, or press Enter to cancel: ').strip()
    if confirm != "RESET DATA":
        print("  Cancelled.")
        return

    # Reset Excel tracking columns
    print("  Resetting Excel stats...", flush=True)
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        for row in ws.iter_rows(min_row=2):
            if row[0].value is None:
                continue
            row[7].value = 0   # # Of times asked
            row[8].value = 0   # RIGHT
            row[9].value = 0   # WRONG
        wb.save(EXCEL_PATH)
        wb.close()
        print("  Excel stats cleared.")
    except PermissionError:
        print("  ERROR: Excel file is open. Close it and try again.")
        return
    except Exception as exc:
        print(f"  ERROR resetting Excel: {exc}")
        return

    # Delete session log
    if os.path.exists(LOG_PATH):
        try:
            os.remove(LOG_PATH)
            print("  Session log deleted.")
        except Exception as exc:
            print(f"  ERROR deleting log: {exc}")
    else:
        print("  No session log found (nothing to delete).")

    print(f"\n  All data reset. Starting fresh!")
    print(DIVIDER)


def choose_mode() -> str:
    """Return 'F', 'S', or 'W'."""
    while True:
        print(f"\n{DIVIDER}")
        print("  SELECT MODE")
        print(DIVIDER)
        print("  F  Full Exam Mode  (75 q, real SIE distribution)")
        print("  S  Study Mode      (choose sections + target count)")
        print("  W  Weak Areas      (questions you've gotten wrong)")
        print("  R  Reset Data      (clear all stats and logs)")
        print("  H  Help            Q  Quit")
        print(THIN_DIV)
        choice = _input("  Your choice: ")
        if choice in ("F", "S", "W", "R"):
            return choice
        if choice == "Q":
            raise SystemExit
        if choice not in ("H", "?"):
            print("  Invalid — please enter F, S, W, or R.")


def choose_main_section() -> str:
    """Return a section key '1'–'4' or 'ALL'."""
    while True:
        print(f"\n{DIVIDER}")
        print("  SELECT SECTION")
        print(DIVIDER)
        for k, v in SECTION_MAP.items():
            print(f"  {k}  Section {k}: {v['name']}")
        print("  A  All Sections")
        print("  H  Help            Q  Back")
        print(THIN_DIV)
        choice = _input("  Your choice: ")
        if choice in SECTION_MAP:
            return choice
        if choice == "A":
            return "ALL"
        if choice == "Q":
            return "Q"
        if choice not in ("H", "?"):
            print("  Invalid — enter a section number (1–4), A, or Q.")


def checklist_menu(items: list, counts: dict, title: str) -> list:
    """
    Interactive toggle checklist. Returns list of selected item strings.
    items  — list of strings (subsection codes)
    counts — dict mapping item -> int (question count for display)
    title  — header text
    """
    selected = set()

    while True:
        print(f"\n{DIVIDER}")
        print(f"  {title}")
        print(DIVIDER)
        for i, item in enumerate(items, 1):
            mark  = "X" if item in selected else " "
            count = counts.get(item, 0)
            # Try to get a friendly name
            name  = ""
            for sk, sv in SECTION_MAP.items():
                if item in sv["sub_names"]:
                    name = sv["sub_names"][item]
                    break
            print(f"  [{mark}] {i}.  {item}  {name:<40}  ({count} q)")
        print(THIN_DIV)
        print("  A  Select All    C  Clear All    Enter  Confirm")
        print("  H  Help          Q  Back")
        print(THIN_DIV)

        raw = input("  Toggle (number/A/C/Enter/Q): ").strip().upper()

        if raw == "":
            if not selected:
                print("  Select at least one subsection first.")
            else:
                return list(selected)
        elif raw == "A":
            selected = set(items)
        elif raw == "C":
            selected = set()
        elif raw == "Q":
            return []
        elif raw in ("H", "?"):
            print(HELP_TEXT)
        else:
            try:
                idx = int(raw) - 1
                if 0 <= idx < len(items):
                    item = items[idx]
                    if item in selected:
                        selected.discard(item)
                    else:
                        selected.add(item)
                else:
                    print(f"  Invalid — enter a number between 1 and {len(items)}.")
            except ValueError:
                print("  Invalid input.")


def get_target_count(available: int, adaptive_on: bool) -> tuple:
    """
    Ask user how many questions and whether adaptive mode is on.
    Returns (target: int, adaptive: bool).
    """
    print(f"\n  Questions available: {available}")
    print(f"  Adaptive learning:  {'ON' if adaptive_on else 'OFF'}  (toggle with A)")

    while True:
        raw = input(f"  How many questions? (Enter = all {available}, A = toggle adaptive): ").strip().upper()
        if raw == "":
            return available, adaptive_on
        if raw == "A":
            adaptive_on = not adaptive_on
            print(f"  Adaptive learning: {'ON' if adaptive_on else 'OFF'}")
            continue
        if raw in ("H", "?"):
            print(HELP_TEXT)
            continue
        try:
            n = int(raw)
            if n <= 0:
                print("  Enter a positive number.")
            else:
                return n, adaptive_on
        except ValueError:
            print("  Invalid — enter a number or press Enter.")

# ---------------------------------------------------------------------------
# Question gathering
# ---------------------------------------------------------------------------

def filter_by_subsections(all_questions: list, subsections: list) -> list:
    return [q for q in all_questions if q["section"] in subsections]


def gather_study_questions(all_questions: list, selected_subs: list,
                           section_key: str, target: int,
                           adaptive: bool) -> list:
    """
    Build a question list for Study Mode.
    Fills from neighboring subsections (same main section) if needed.
    """
    pool = filter_by_subsections(all_questions, selected_subs)

    if len(pool) < target and section_key != "ALL":
        # Pull from other subsections in the same main section
        all_subs = SECTION_MAP[section_key]["subsections"]
        neighbors = [s for s in all_subs if s not in selected_subs]
        for neighbor in neighbors:
            if len(pool) >= target:
                break
            extra = filter_by_subsections(all_questions, [neighbor])
            pool += extra
        if len(pool) < target:
            print(f"\n  [Note] Only {len(pool)} questions available in Section {section_key} — using all of them.")

    if adaptive:
        return weighted_sample(pool, target)
    else:
        sampled = pool[:] if target >= len(pool) else random.sample(pool, target)
        random.shuffle(sampled)
        return sampled


def gather_exam_questions(all_questions: list, adaptive: bool) -> list:
    """Build a 75-question list matching the real SIE exam distribution."""
    result = []
    for sk, count in EXAM_DISTRIBUTION.items():
        section_subs = SECTION_MAP[sk]["subsections"]
        pool = filter_by_subsections(all_questions, section_subs)
        if len(pool) == 0:
            print(f"  [Warning] Section {sk} has no questions — skipping its {count} slots.")
            continue
        if len(pool) < count:
            print(f"  [Note] Section {sk} only has {len(pool)} questions (need {count}) — using all available.")

        if adaptive:
            sampled = weighted_sample(pool, count)
        else:
            sampled = pool[:] if count >= len(pool) else random.sample(pool, count)

        result.extend(sampled)

    random.shuffle(result)
    return result

# ---------------------------------------------------------------------------
# Quiz engine
# ---------------------------------------------------------------------------

def run_quiz(questions: list) -> tuple:
    """
    Run the quiz loop. Returns (correct_count, total, section_results).
    section_results: dict mapping main section key -> (correct, total)
    """
    correct_count = 0
    total         = len(questions)
    # Track per-section for Full Exam breakdown
    section_results = {}  # "1" -> [correct, total]

    for i, q in enumerate(questions, 1):
        print(f"\n{DIVIDER}")
        print(f"  Question {i} of {total}   |   Section {q['section']}")
        print(DIVIDER)

        # Wrap long question text at ~60 chars
        words = q["question"].split()
        line, lines = "", []
        for w in words:
            if len(line) + len(w) + 1 > 60:
                lines.append(line)
                line = w
            else:
                line = (line + " " + w).strip()
        if line:
            lines.append(line)
        for ln in lines:
            print(f"  {ln}")
        print()

        # Shuffle answer choices
        letters = ["A", "B", "C", "D"]
        random.shuffle(letters)
        mapping = {str(num): letter for num, letter in enumerate(letters, 1)}

        for num, letter in mapping.items():
            print(f"    {num}. {q['choices'][letter]}")
        print()

        # Get answer
        answer = None
        while answer is None:
            raw = _input("  Your answer (1–4, H=help, Q=quit): ")
            if raw == "Q":
                print("\n  Returning to main menu...")
                # Save partial progress
                return correct_count, i - 1, section_results
            if raw in ("1", "2", "3", "4"):
                answer = raw
            elif raw not in ("H", "?"):
                print("  Invalid — enter 1, 2, 3, or 4.")

        chosen_letter  = mapping[answer]
        is_correct     = (chosen_letter == q["correct"])
        correct_letter = q["correct"]
        correct_text   = q["choices"][correct_letter]

        if is_correct:
            correct_count += 1
            print(f"\n  CORRECT!  {correct_text}")
        else:
            chosen_text = q["choices"][chosen_letter]
            print(f"\n  INCORRECT.  You chose:     {chosen_text}")
            print(f"  Correct answer:  {correct_text}")

        # Track per-section
        sec_key = q["section"].split(".")[0]
        if sec_key not in section_results:
            section_results[sec_key] = [0, 0]
        section_results[sec_key][1] += 1
        if is_correct:
            section_results[sec_key][0] += 1

        # Save to Excel
        save_result(q["row"], is_correct)
        print(THIN_DIV)

    return correct_count, total, section_results


def show_score(correct: int, total: int,
               section_results: dict = None, exam_mode: bool = False) -> None:
    """Print final score. Optionally show per-section breakdown for exam mode."""
    print(f"\n{DIVIDER}")
    print("  RESULTS")
    print(DIVIDER)

    if total == 0:
        print("  No questions were answered.")
        print(DIVIDER)
        return

    pct = correct / total * 100

    if exam_mode and section_results:
        for sk in sorted(SECTION_MAP.keys()):
            if sk in section_results:
                sc, st = section_results[sk]
                sp = sc / st * 100 if st > 0 else 0
                name = SECTION_MAP[sk]["name"]
                print(f"  Section {sk}: {sc:>2}/{st:<2}  ({sp:.1f}%)  {name}")
        print(THIN_DIV)

    print(f"  Overall:  {correct} / {total}  ({pct:.1f}%)")

    if exam_mode:
        result = "PASS" if pct >= PASS_THRESHOLD * 100 else "FAIL"
        print(f"\n  SIE Pass threshold: {PASS_THRESHOLD*100:.0f}%   Result: {result}")

    print(DIVIDER)

# ---------------------------------------------------------------------------
# Session log
# ---------------------------------------------------------------------------

def _sections_label(mode: str, selected_subs: list) -> str:
    if mode == "F":
        return "Full Exam"
    if mode == "W":
        return "Weak Areas"
    return "+".join(sorted(selected_subs))

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    # Fix Windows terminal Unicode issues
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    print(f"\n{DIVIDER}")
    print("  SIE EXAM QUIZZER")
    print("  Type H at any prompt for help.")
    print(DIVIDER)

    # Load question bank once
    print("  Loading question bank...", flush=True)
    all_questions = load_all_questions()
    if not all_questions:
        print("\n  No questions found in the Excel file. Add questions and re-run.")
        sys.exit(0)

    # Show progress dashboard
    show_dashboard(all_questions)

    adaptive_default = True  # adaptive learning on by default

    try:
        while True:
            mode = choose_mode()
            selected_subs = []

            # Reset — clear stats then reload so dashboard shows zeroed counts
            if mode == "R":
                reset_data()
                all_questions = load_all_questions()
                show_dashboard(all_questions)
                continue

            # --- Full Exam Mode ---
            if mode == "F":
                print(f"\n  Building Full Exam (75 questions)...")
                questions = gather_exam_questions(all_questions, adaptive=True)
                if not questions:
                    print("  No questions available for Full Exam Mode. Add questions and re-run.")
                    continue
                print(f"  {len(questions)} questions loaded. Good luck!\n")
                correct, total, section_results = run_quiz(questions)
                show_score(correct, total, section_results, exam_mode=True)
                append_session_log("Full Exam", "All", 75, correct, total)

            # --- Weak Areas Mode ---
            elif mode == "W":
                weak = [q for q in all_questions if q["wrong"] > 0]
                if not weak:
                    print("\n  No wrong answers yet — keep studying!")
                    continue
                print(f"\n  {len(weak)} weak-area questions found.")
                target, _ = get_target_count(len(weak), adaptive_default)
                sampled = random.sample(weak, min(target, len(weak)))
                random.shuffle(sampled)
                correct, total, section_results = run_quiz(sampled)
                show_score(correct, total)
                append_session_log("Weak Areas", "Weak", target, correct, total)

            # --- Study Mode ---
            else:
                sec_key = choose_main_section()
                if sec_key == "Q":
                    continue

                if sec_key == "ALL":
                    all_subs = []
                    for sv in SECTION_MAP.values():
                        all_subs.extend(sv["subsections"])
                    counts = {}
                    for sub in all_subs:
                        counts[sub] = sum(1 for q in all_questions if q["section"] == sub)
                    selected_subs = checklist_menu(all_subs, counts, "SELECT SUBSECTIONS (All Sections)")
                    effective_section = "ALL"
                else:
                    subs = SECTION_MAP[sec_key]["subsections"]
                    counts = {}
                    for sub in subs:
                        counts[sub] = sum(1 for q in all_questions if q["section"] == sub)
                    section_title = f"SELECT SUBSECTIONS — Section {sec_key}"
                    selected_subs = checklist_menu(subs, counts, section_title)
                    effective_section = sec_key

                if not selected_subs:
                    continue

                pool_size = len(filter_by_subsections(all_questions, selected_subs))
                # Include neighbors in the available count hint
                if effective_section != "ALL":
                    all_sec_subs = SECTION_MAP[effective_section]["subsections"]
                    total_in_section = sum(
                        1 for q in all_questions
                        if q["section"] in all_sec_subs
                    )
                    if total_in_section > pool_size:
                        print(f"\n  ({total_in_section - pool_size} more questions available from "
                              f"other subsections in Section {effective_section} for auto-fill)")

                target, adaptive_default = get_target_count(pool_size, adaptive_default)
                questions = gather_study_questions(
                    all_questions, selected_subs, effective_section,
                    target, adaptive_default
                )

                if not questions:
                    print("  No questions found for that selection. Try a different section.")
                    continue

                correct, total, section_results = run_quiz(questions)
                show_score(correct, total)
                append_session_log(
                    "Study", _sections_label(mode, selected_subs),
                    target, correct, total
                )

            # Continue prompt
            print()
            again = _input("  Study another session? (Y / N): ")
            if again != "Y":
                break

            # Reload questions so updated stats are reflected in next session
            all_questions = load_all_questions()
            show_dashboard(all_questions)

    except (KeyboardInterrupt, SystemExit):
        pass

    print("\n  Goodbye! Good luck on the SIE.\n")


if __name__ == "__main__":
    main()
