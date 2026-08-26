from __future__ import annotations

import math
from pathlib import Path
from textwrap import wrap

from PIL import Image, ImageDraw, ImageFont


NAVY = "#0F172A"
MUTED = "#475569"
BORDER = "#CBD5E1"


def font(size: int, bold: bool = False):
    choices = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for path in choices:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_wrapped(draw, xy, text, size=18, bold=False, fill=MUTED, width=48, line_gap=5):
    x, y = xy
    face = font(size, bold)
    for line in wrap(str(text), width=width) or [""]:
        draw.text((x, y), line, font=face, fill=fill)
        y += size + line_gap
    return y


def draw_box(draw, xy, title, lines, fill="#FFFFFF", accent="#2563EB"):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle(xy, radius=20, fill=fill, outline="#94A3B8", width=3)
    draw.rectangle((x1, y1, x1 + 9, y2), fill=accent)
    draw.text((x1 + 28, y1 + 20), title, font=font(25, True), fill=NAVY)
    y = y1 + 64
    for line in lines:
        draw.text((x1 + 28, y), str(line), font=font(18), fill=MUTED)
        y += 29


def arrow(draw, start, end, label=""):
    x1, y1 = start
    x2, y2 = end
    draw.line((x1, y1, x2, y2), fill="#64748B", width=4)
    angle = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - 18 * math.cos(angle - 0.55), y2 - 18 * math.sin(angle - 0.55))
    p2 = (x2 - 18 * math.cos(angle + 0.55), y2 - 18 * math.sin(angle + 0.55))
    draw.polygon([(x2, y2), p1, p2], fill="#64748B")
    if label:
        mx, my = (x1 + x2) // 2, (y1 + y2) // 2
        box = draw.textbbox((0, 0), label, font=font(15, True))
        width = box[2] - box[0]
        draw.rounded_rectangle((mx - width // 2 - 8, my - 25, mx + width // 2 + 8, my + 2), radius=6, fill="#F8FAFC")
        draw.text((mx - width // 2, my - 22), label, font=font(15, True), fill=MUTED)


def create_vibe_layout(path: Path):
    image = Image.new("RGB", (1900, 1120), "#E2E8F0")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((45, 35, 1855, 1080), radius=24, fill="#FFFFFF", outline="#94A3B8", width=4)
    draw.rectangle((45, 35, 1855, 115), fill=NAVY)
    draw.text((80, 61), "Vibe Code  /  Project: Site Inspection Tracker", font=font(27, True), fill="#FFFFFF")
    draw.rounded_rectangle((1285, 55, 1455, 98), radius=10, fill="#0F766E")
    draw.text((1320, 65), "READY", font=font(20, True), fill="#FFFFFF")
    for x, label in [(1480, "History"), (1650, "Full preview")]:
        draw.rounded_rectangle((x, 52, x + 165, 101), radius=10, fill="#F8FAFC")
        draw.text((x + 22, 65), label, font=font(18, True), fill=NAVY)

    draw.rectangle((45, 115, 1855, 183), fill="#F8FAFC")
    metadata = ["Owner: you", "Version: 4", "Workspace: Active", "Last saved: 10:42", "Usage cap: within limit"]
    x = 78
    for item in metadata:
        draw.text((x, 138), item, font=font(17, True), fill=MUTED)
        x += 335

    draw.rounded_rectangle((70, 210, 890, 1000), radius=18, fill="#F8FAFC", outline=BORDER, width=3)
    draw.text((100, 238), "Conversation and change control", font=font(24, True), fill=NAVY)
    draw.rounded_rectangle((105, 300, 665, 385), radius=16, fill="#DBEAFE")
    draw_wrapped(draw, (135, 320), "Add a weekly safety inspection dashboard with overdue actions.", 19, fill="#1E3A8A", width=44)
    draw.rounded_rectangle((230, 420, 845, 610), radius=16, fill="#ECFEFF")
    draw.text((260, 443), "Understanding → Building → Checking", font=font(20, True), fill="#0F766E")
    draw_wrapped(draw, (260, 485), "I added the dashboard, API endpoint and database migration. Review the safe diff and preview before accepting.", 18, fill="#334155", width=55)
    for x, label, color in [(260, "Review diff", "#2563EB"), (430, "Accept", "#0F766E"), (565, "Reject", "#B91C1C"), (700, "Iterate", "#7C3AED")]:
        draw.rounded_rectangle((x, 545, x + 125, 590), radius=9, fill=color)
        draw.text((x + 14, 557), label, font=font(16, True), fill="#FFFFFF")
    draw.rounded_rectangle((105, 650, 845, 810), radius=16, fill="#FFFFFF", outline=BORDER, width=2)
    draw.text((132, 672), "Describe the next business outcome…", font=font(18), fill="#64748B")
    draw.rounded_rectangle((690, 750, 815, 792), radius=9, fill="#2563EB")
    draw.text((731, 760), "Send", font=font(16, True), fill="#FFFFFF")
    for x, label in [(105, "Stop run"), (245, "Save version"), (415, "Restore")]:
        draw.rounded_rectangle((x, 835, x + 135, 880), radius=9, fill="#FFFFFF", outline="#94A3B8", width=2)
        draw.text((x + 14, 847), label, font=font(16, True), fill="#334155")
    draw.rounded_rectangle((105, 910, 845, 968), radius=12, fill="#FEF3C7")
    draw.text((135, 927), "Status: Ready for review  •  autosave is recovery only", font=font(17, True), fill="#92400E")

    draw.rounded_rectangle((920, 210, 1825, 1000), radius=18, fill="#FFFFFF", outline=BORDER, width=3)
    draw.text((950, 238), "Authenticated live preview", font=font(24, True), fill=NAVY)
    draw.rounded_rectangle((950, 295, 1795, 355), radius=10, fill="#F1F5F9")
    draw.text((980, 314), "Owner-only preview  •  generation 12  •  live update connected", font=font(17, True), fill=MUTED)
    draw.rounded_rectangle((950, 390, 1795, 925), radius=14, fill="#F8FAFC", outline=BORDER, width=2)
    draw.text((990, 430), "Weekly Safety Dashboard", font=font(29, True), fill=NAVY)
    for x, value, label, color in [(990, "18", "Open actions", "#2563EB"), (1240, "5", "Overdue", "#B91C1C"), (1490, "92%", "Completion", "#0F766E")]:
        draw.rounded_rectangle((x, 505, x + 210, 635), radius=14, fill="#FFFFFF", outline=color, width=3)
        draw.text((x + 25, 528), value, font=font(34, True), fill=color)
        draw.text((x + 25, 582), label, font=font(17), fill=MUTED)
    draw.rounded_rectangle((990, 680, 1750, 860), radius=12, fill="#FFFFFF", outline=BORDER, width=2)
    draw.text((1020, 705), "Recent inspections", font=font(20, True), fill=NAVY)
    for index, row in enumerate(["Tower A  •  Complete", "Basement  •  2 overdue", "Podium  •  In progress"]):
        draw.text((1020, 755 + index * 36), row, font=font(17), fill=MUTED)
    draw.text((70, 1040), "Product boundary: no source editor, terminal, infrastructure controls, arbitrary capability install or Deploy button.", font=font(19, True), fill="#991B1B")
    image.save(path)


def create_use_case_flow(path: Path, use_case: tuple[str, ...]):
    uid, title, actors, module, layer, precondition, flow, rules, security, failure, evidence, stories = use_case
    steps = [step.strip() for step in flow.replace("\n", " ").split("→") if step.strip()]
    if len(steps) > 5:
        steps = steps[:4] + [" / ".join(steps[4:])]
    image = Image.new("RGB", (1900, 900), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    draw.text((55, 30), f"{uid} — {title}", font=font(36, True), fill=NAVY)
    draw.text((55, 82), f"Actor: {actors.replace(chr(10), ' / ')}  •  Layer: {layer}", font=font(19), fill=MUTED)
    draw.rounded_rectangle((55, 125, 1845, 220), radius=14, fill="#ECFEFF", outline="#5EEAD4", width=2)
    draw.text((82, 145), "Precondition", font=font(19, True), fill="#0F766E")
    draw_wrapped(draw, (230, 145), precondition, 17, fill="#334155", width=145)
    gap = 55
    start_x = 70
    usable = 1900 - 2 * start_x
    box_w = min(310, int((usable - gap * (len(steps) - 1)) / max(len(steps), 1)))
    for index, step in enumerate(steps):
        x1 = start_x + index * (box_w + gap)
        x2 = x1 + box_w
        draw_box(draw, (x1, 275, x2, 460), str(index + 1), wrap(step, width=max(18, box_w // 11)), "#FFFFFF", "#2563EB")
        if index < len(steps) - 1:
            arrow(draw, (x2, 367), (x2 + gap, 367))
    draw.rounded_rectangle((55, 520, 895, 680), radius=16, fill="#FFF7ED", outline="#FDBA74", width=2)
    draw.text((82, 542), "Alternative / failure path", font=font(20, True), fill="#B45309")
    draw_wrapped(draw, (82, 582), failure, 17, fill=MUTED, width=78)
    draw.rounded_rectangle((930, 520, 1845, 680), radius=16, fill="#F0FDF4", outline="#86EFAC", width=2)
    draw.text((957, 542), "Postcondition / evidence", font=font(20, True), fill="#15803D")
    draw_wrapped(draw, (957, 582), evidence, 17, fill=MUTED, width=83)
    draw.rounded_rectangle((55, 720, 1845, 845), radius=16, fill="#F5F3FF", outline="#C4B5FD", width=2)
    draw.text((82, 742), f"Business Rules: {rules}", font=font(18, True), fill="#6D28D9")
    draw_wrapped(draw, (82, 779), f"Security guard: {security}  •  Stories: {stories}  •  Module: {module}", 16, fill=MUTED, width=165)
    image.save(path)


def create_state_diagram(path: Path, machine: dict):
    outgoing = {state["name"]: [] for state in machine["states"]}
    for edge in machine["transitions"]:
        outgoing[edge["from"]].append(edge["to"])
    row_height = 92
    height = 300 + row_height * len(machine["states"])
    image = Image.new("RGB", (1900, height), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    draw.text((55, 28), machine["field"], font=font(36, True), fill=NAVY)
    draw.text((55, 78), f"Owner: {machine['owner']}", font=font(19), fill=MUTED)
    draw.rounded_rectangle((55, 118, 1845, 210), radius=14, fill="#FEF3C7", outline="#F59E0B", width=2)
    draw.text((82, 137), "Guard", font=font(20, True), fill="#92400E")
    draw_wrapped(draw, (175, 137), machine["guard_summary"], 18, fill=MUTED, width=135)
    draw.text((75, 242), "FROM STATE", font=font(18, True), fill=MUTED)
    draw.text((650, 242), "LEGAL DIRECTED TRANSITIONS", font=font(18, True), fill=MUTED)
    for index, state in enumerate(machine["states"]):
        y = 280 + index * row_height
        name = state["name"]
        initial = bool(state.get("initial", name == machine.get("initial")))
        terminal = bool(state.get("terminal", name in machine.get("terminal", [])))
        fill = "#DBEAFE" if initial else ("#DCFCE7" if terminal else "#FFFFFF")
        accent = "#2563EB" if initial else ("#15803D" if terminal else "#64748B")
        draw.rounded_rectangle((65, y, 535, y + 66), radius=12, fill=fill, outline=accent, width=3)
        marker = "● " if initial else ("◎ " if terminal else "")
        draw.text((90, y + 13), marker + name, font=font(18, True), fill=NAVY)
        draw.text((350, y + 19), state.get("ui", state.get("normal_user_mapping", "")), font=font(14), fill=MUTED)
        arrow(draw, (535, y + 33), (625, y + 33))
        targets = outgoing[name]
        target_text = "  •  ".join(targets) if targets else "TERMINAL — no outgoing transition"
        draw.rounded_rectangle((625, y, 1835, y + 66), radius=12, fill="#FFFFFF", outline=BORDER, width=2)
        draw_wrapped(draw, (650, y + 12), target_text, 16, terminal, "#15803D" if terminal else "#334155", width=125, line_gap=3)
    draw.text((65, height - 20), f"{len(machine['transitions'])} legal edges. Any unlisted or unknown transition fails closed and is audited.", font=font(16, True), fill="#991B1B")
    image.save(path)


def create_sequence_diagram(path: Path, title: str, lanes: list[str], steps: list[tuple[int, int, str]], guard: str):
    image = Image.new("RGB", (1900, 1020), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    draw.text((55, 28), title, font=font(36, True), fill=NAVY)
    xs = [110 + index * int(1680 / max(1, len(lanes) - 1)) for index in range(len(lanes))]
    for x, lane in zip(xs, lanes):
        box = draw.textbbox((0, 0), lane, font=font(18, True))
        draw.text((x - (box[2] - box[0]) // 2, 105), lane, font=font(18, True), fill=NAVY)
        draw.line((x, 150, x, 850), fill=BORDER, width=3)
    y = 190
    for index, (source, target, label) in enumerate(steps, 1):
        arrow(draw, (xs[source], y), (xs[target], y), f"{index}  {label}")
        y += 82
    draw.rounded_rectangle((70, 885, 1830, 980), radius=14, fill="#FEF3C7", outline="#F59E0B", width=2)
    draw.text((96, 905), "Fail-closed guard", font=font(19, True), fill="#92400E")
    draw_wrapped(draw, (285, 905), guard, 17, fill=MUTED, width=135)
    image.save(path)


def create_ui_component_contract(path: Path):
    image = Image.new("RGB", (1900, 980), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    draw.text((55, 28), "Vibe Page — Component and Trust-Boundary Contract", font=font(36, True), fill=NAVY)
    draw_box(draw, (65, 145, 480, 350), "Product shell", ["Project metadata", "status + version", "History / full preview"], "#EFF6FF", "#2563EB")
    draw_box(draw, (65, 430, 480, 720), "Conversation pane", ["durable messages/events", "Stop / Save / Restore", "diff decision controls", "safe usage evidence"], "#ECFEFF", "#0F766E")
    draw_box(draw, (620, 250, 1060, 605), "NestJS control plane", ["Prisma + PostgreSQL", "Redis server-side session", "owner auth + CSRF", "command idempotency + CAS"], "#FFFFFF", "#64748B")
    draw_box(draw, (1200, 145, 1815, 350), "Preview Gateway", ["one-time short-lived JWT", "host-only preview cookie", "owner/session/project/generation", "strip platform credentials"], "#FFF7ED", "#D97706")
    draw_box(draw, (1200, 430, 1815, 720), "Generated preview", ["Vite frontend", "FastAPI /api", "HMR over authenticated WSS", "zero or one mutable replica"], "#F0FDF4", "#16A34A")
    arrow(draw, (480, 245), (620, 325), "product API")
    arrow(draw, (480, 570), (620, 500), "chat / commands")
    arrow(draw, (1060, 325), (1200, 245), "ticket / route")
    arrow(draw, (1505, 350), (1505, 430), "fixed targets")
    arrow(draw, (1200, 575), (1060, 530), "status / evidence")
    draw.rounded_rectangle((100, 790, 1800, 925), radius=18, fill="#FEE2E2", outline="#FCA5A5", width=2)
    draw.text((130, 815), "Forbidden across the boundary", font=font(21, True), fill="#991B1B")
    draw_wrapped(draw, (130, 858), "No source editor, terminal, raw commands, infrastructure manifests, provider credentials, domain-wide platform cookies, arbitrary upstreams or Stage 1 Deploy control.", 18, fill=MUTED, width=150)
    image.save(path)
