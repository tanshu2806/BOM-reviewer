"""Generate fictional, presentation-ready drawing/BOM pairs for the prototype."""

import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


OUTPUT_DIR = Path(__file__).resolve().parents[1] / "demo" / "prototype_samples"
PAGE_W, PAGE_H = landscape(A3)
INK = colors.HexColor("#183047")
BLUE = colors.HexColor("#1f587c")
PALE_BLUE = colors.HexColor("#eaf1f6")
GREY = colors.HexColor("#657786")
GRID = colors.HexColor("#9aabb7")
ORANGE = colors.HexColor("#d97706")
RED = colors.HexColor("#c2413b")


def item(item_id, description, material, size, qty, partno):
    return {
        "item": str(item_id),
        "description": description,
        "material": material,
        "size": size,
        "qty": qty,
        "partno": partno,
        "remarks": "",
    }


SAMPLES = [
    {
        "stem": "01_process_pump_skid",
        "title": "HORIZONTAL PROCESS PUMP SKID",
        "subtitle": "General arrangement and fabrication bill of materials",
        "drawing_no": "NB-PS-2401-GA-001",
        "revision": "B",
        "equipment": "P-2401 A/B  |  PROCESS TRANSFER PUMP PACKAGE",
        "kind": "pump",
        "items": [
            item(1, "BASEPLATE, COMMON MACHINED", "ASTM A36", "1800 x 620 x 16", 1, "BP-2401-01"),
            item(2, "CENTRIFUGAL PUMP, 50-32-200", "ASTM A216 WCB", "DN50 x DN32", 1, "PU-2401-01"),
            item(3, "ELECTRIC MOTOR, 7.5 kW, 2-POLE", "IEC 60034", "400 V / 50 Hz", 1, "MO-2401-01"),
            item(4, "FLEXIBLE GRID COUPLING", "AISI 1045", "SIZE 1070", 1, "CP-2401-01"),
            item(5, "COUPLING GUARD, REMOVABLE", "SS 304", "2.0 mm THK", 1, "CG-2401-01"),
            item(6, "SUCTION ISOLATION VALVE, GATE", "ASTM A216 WCB", "DN50 / CLASS 150", 1, "XV-2401-01"),
            item(7, "DISCHARGE CHECK VALVE, SWING", "ASTM A216 WCB", "DN32 / CLASS 150", 1, "NRV-2401-01"),
            item(8, "DISCHARGE ISOLATION VALVE, GATE", "ASTM A216 WCB", "DN32 / CLASS 150", 1, "XV-2401-02"),
            item(9, "SUCTION PRESSURE GAUGE, 0-10 bar", "SS 316", "DN15 / 100 mm DIAL", 1, "PI-2401-01"),
            item(10, "DISCHARGE PRESSURE GAUGE, 0-25 bar", "SS 316", "DN15 / 100 mm DIAL", 1, "PI-2401-02"),
            item(11, "ECCENTRIC SUCTION REDUCER, FLAT TOP", "ASTM A234 WPB", "DN50 x DN32", 1, "RD-2401-01"),
            item(12, "DISCHARGE PIPE SPOOL, FLANGED", "ASTM A106 GR B", "DN32 / SCH 40", 2, "SP-2401-01"),
            item(13, "PIPE SUPPORT, CLAMP TYPE", "ASTM A36", "DN32", 2, "PS-2401-01"),
            item(14, "DRAIN VALVE, BALL TYPE", "SS 316", "DN15 / CLASS 800", 1, "DV-2401-01"),
            item(15, "HOLD-DOWN BOLT SET WITH NUTS", "ASTM F1554 GR 55", "M20 x 180", 4, "HD-2401-01"),
            item(16, "EARTHING LUG, ATTACHED", "SS 304", "50 x 40 x 6", 2, "EL-2401-01"),
            item(17, "NAMEPLATE, EQUIPMENT DATA", "SS 316", "100 x 70", 1, "NP-2401-01"),
        ],
        "bom_changes": {},
        "missing_bom": [],
        "extra_bom": [],
        "notes": [
            "Pump and motor to be aligned on common baseplate after grouting.",
            "All pressure connections are flanged to ASME B16.5, Class 150 unless noted.",
            "Provide 3 mm drain holes at low points; protect machined surfaces for shipment.",
        ],
    },
    {
        "stem": "02_shell_tube_exchanger",
        "title": "SHELL-AND-TUBE HEAT EXCHANGER",
        "subtitle": "Two-pass bundle arrangement and exchanger bill of materials",
        "drawing_no": "NB-HX-3102-GA-014",
        "revision": "C",
        "equipment": "E-3102  |  FIXED TUBESHEET / HORIZONTAL MOUNTING",
        "kind": "exchanger",
        "items": [
            item(1, "STATIONARY HEAD, CHANNEL", "ASTM A516 GR 70", "DN600 / 18 THK", 1, "CH-3102-01"),
            item(2, "STATIONARY HEAD, BONNET", "ASTM A516 GR 70", "DN600 / 18 THK", 1, "BH-3102-01"),
            item(3, "STATIONARY HEAD FLANGE", "ASTM A105", "DN600 / CLASS 300", 2, "FL-3102-01"),
            item(4, "CHANNEL COVER, BOLTED", "ASTM A516 GR 70", "DN600 / 22 THK", 1, "CV-3102-01"),
            item(5, "STATIONARY HEAD NOZZLE", "ASTM A106 GR B", "DN100 / SCH 40", 2, "NZ-3102-01"),
            item(6, "STATIONARY TUBESHEET, DRILLED", "ASTM A266 GR 2", "OD 600 x 55 THK", 1, "TS-3102-01"),
            item(7, "HEAT TRANSFER TUBES, SEAMLESS", "ASTM A213 TP316L", "19.05 OD x 2.11 THK x 3200 LG", 96, "TB-3102-01"),
            item(8, "SHELL, LONGITUDINAL SEAM", "ASTM A516 GR 70", "ID 610 x 12 THK", 1, "SH-3102-01"),
            item(9, "SHELL COVER, FLOATING END", "ASTM A516 GR 70", "ID 580 x 12 THK", 1, "FC-3102-01"),
            item(10, "SHELL FLANGE, STATIONARY END", "ASTM A105", "DN650 / CLASS 300", 1, "FL-3102-02"),
            item(11, "SHELL FLANGE, REAR HEAD END", "ASTM A105", "DN650 / CLASS 300", 1, "FL-3102-03"),
            item(12, "SHELL NOZZLE, FLANGED", "ASTM A105", "DN150 / CLASS 300", 2, "NZ-3102-02"),
            item(13, "SHELL COVER FLANGE", "ASTM A105", "DN580 / CLASS 300", 1, "FL-3102-04"),
            item(14, "SHELL EXPANSION JOINT", "ASTM A240 TP316L", "ID 610 / 3 CONVOLUTIONS", 1, "EJ-3102-01"),
            item(15, "FLOATING TUBESHEET, DRILLED", "ASTM A266 GR 2", "OD 545 x 48 THK", 1, "FT-3102-01"),
            item(16, "FLOATING HEAD COVER", "ASTM A516 GR 70", "DN500 / 20 THK", 1, "FC-3102-02"),
            item(17, "FLOATING HEAD FLANGE", "ASTM A105", "DN500 / CLASS 300", 1, "FL-3102-05"),
            item(18, "FLOATING HEAD BACKING DEVICE", "ASTM A516 GR 70", "DN500 / 16 THK", 1, "BD-3102-01"),
            item(19, "FLOATING HEAD NOZZLE", "ASTM A106 GR B", "DN100 / SCH 40", 1, "NZ-3102-03"),
            item(20, "FLOATING HEAD COVER, EXTERNAL", "ASTM A516 GR 70", "OD 680 x 16 THK", 1, "EC-3102-01"),
            item(21, "FLOATING TUBESHEET SKIRT", "ASTM A240 TP316L", "OD 548 x 8 THK", 1, "SK-3102-01"),
            item(22, "FLOATING HEAD PACKING, INNER", "GRAPHITE / SS 316", "12 x 12 SECTION", 4, "PK-3102-01"),
            item(23, "PACKING BOX FLANGE", "ASTM A105", "DN500 / CLASS 300", 1, "FL-3102-06"),
            item(24, "FLOATING HEAD PACKING, OUTER", "GRAPHITE / SS 316", "12 x 12 SECTION", 4, "PK-3102-02"),
            item(25, "PACKING GLAND, FIXED", "ASTM A105", "DN500 / 16 THK", 1, "PG-3102-01"),
            item(26, "PACKING GLAND, FOLLOWER", "ASTM A105", "DN500 / 16 THK", 1, "PG-3102-02"),
            item(27, "LANTERN RING", "ASTM A240 TP316L", "DN500 / 18 THK", 1, "LR-3102-01"),
            item(28, "TIE RODS WITH SPACERS", "ASTM A193 B8M", "M16 x 980", 4, "TR-3102-01"),
            item(29, "TRANSVERSE BAFFLES / SUPPORT PLATES", "ASTM A240 TP316L", "6 THK / 25% CUT", 8, "BF-3102-01"),
            item(30, "SHELL-SIDE IMPINGEMENT PLATE", "ASTM A240 TP316L", "8 THK / 260 x 220", 1, "IP-3102-01"),
            item(31, "LONGITUDINAL BAFFLE", "ASTM A240 TP316L", "6 THK / 520 x 3200", 1, "LB-3102-01"),
            item(32, "TUBE-SIDE PASS PARTITION", "ASTM A240 TP316L", "8 THK / 32 WIDE", 1, "PP-3102-01"),
            item(33, "VENT CONNECTION, FLANGED", "ASTM A105", "DN25 / CLASS 300", 2, "NZ-3102-04"),
            item(34, "DRAIN CONNECTION, FLANGED", "ASTM A105", "DN40 / CLASS 300", 2, "NZ-3102-05"),
            item(35, "INSTRUMENT CONNECTION, COUPLING", "ASTM A105", "DN15 / CLASS 3000", 2, "IC-3102-01"),
            item(36, "SHELL SUPPORT SADDLE", "ASTM A36", "610 OD / 10 THK", 2, "SD-3102-01"),
            item(37, "SHELL LIFTING LUG", "ASTM A36", "PL 20 THK", 4, "LG-3102-01"),
            item(38, "FLOATING HEAD WEAR RING", "ASTM A240 TP316L", "OD 548 x 8 THK", 1, "WR-3102-01"),
        ],
        "bom_changes": {
            "7": {"qty": 92},
            "33": {"material": "ASTM A182 F316L"},
        },
        "missing_bom": ["16"],
        "extra_bom": [],
        "notes": [
            "Tube layout: 96 tubes on 25.4 mm triangular pitch; two tube-side passes.",
            "Shell-side baffles equally spaced; first and last spacing 180 mm nominal.",
            "Gaskets are prototype selections only; confirm process compatibility before use.",
        ],
    },
    {
        "stem": "03_vertical_air_receiver",
        "title": "VERTICAL COMPRESSED-AIR RECEIVER",
        "subtitle": "Vessel arrangement, nozzle orientation and fabrication bill of materials",
        "drawing_no": "NB-AR-0805-GA-003",
        "revision": "A",
        "equipment": "V-0805  |  4.0 m3 NOMINAL / VERTICAL INSTALLATION",
        "kind": "receiver",
        "items": [
            item(1, "VESSEL SHELL COURSE, ROLLED", "ASTM A516 GR 70", "ID 1500 x 12 THK", 1, "VS-0805-01"),
            item(2, "TOP TORISPHERICAL HEAD", "ASTM A516 GR 70", "ID 1500 x 14 THK", 1, "HD-0805-01"),
            item(3, "BOTTOM TORISPHERICAL HEAD", "ASTM A516 GR 70", "ID 1500 x 14 THK", 1, "HD-0805-02"),
            item(4, "SUPPORT SKIRT, CYLINDRICAL", "ASTM A36", "OD 1512 x 8 THK", 1, "SK-0805-01"),
            item(5, "BASE RING, ANCHOR BOLT TYPE", "ASTM A36", "OD 1780 x 25 THK", 1, "BR-0805-01"),
            item(6, "INLET NOZZLE, REINFORCED", "ASTM A105", "DN80 / CLASS 150", 1, "NZ-0805-01"),
            item(7, "OUTLET NOZZLE, REINFORCED", "ASTM A105", "DN80 / CLASS 150", 1, "NZ-0805-02"),
            item(8, "SAFETY RELIEF VALVE, SET 10 bar", "ASTM A216 WCB", "DN25 x DN40", 1, "PSV-0805-01"),
            item(9, "PRESSURE GAUGE, 0-16 bar", "SS 316", "DN15 / 100 mm DIAL", 1, "PI-0805-01"),
            item(10, "PRESSURE TRANSMITTER, 4-20 mA", "SS 316", "0-16 bar / 1/2 NPT", 1, "PT-0805-01"),
            item(11, "AUTO DRAIN VALVE, TIMER TYPE", "BRASS / SS 304", "DN15 / 230 VAC", 1, "DV-0805-01"),
            item(12, "MANUAL DRAIN VALVE, BALL TYPE", "SS 316", "DN20 / CLASS 800", 1, "DV-0805-02"),
            item(13, "MANWAY NECK AND BLIND COVER", "ASTM A516 GR 70", "DN450 / CLASS 150", 1, "MW-0805-01"),
            item(14, "MANWAY GASKET, NON-ASBESTOS", "ARAMID / NBR", "DN450", 1, "GS-0805-01"),
            item(15, "NAMEPLATE BRACKET, ATTACHED", "SS 304", "75 x 50 x 6", 1, "NB-0805-01"),
            item(16, "EARTHING LUG, SHELL", "SS 304", "50 x 40 x 6", 2, "EL-0805-01"),
        ],
        "bom_changes": {
            "6": {"qty": 2},
        },
        "missing_bom": ["14"],
        "extra_bom": [
            item("SP-01", "SPARE PRESSURE GAUGE", "SS 316", "0-16 bar", 1, "SP-0805-01"),
        ],
        "notes": [
            "Design pressure 10 bar(g); hydrostatic test pressure to approved vessel code calculations.",
            "Provide 450 mm clear access in front of manway cover for maintenance.",
            "All nozzle azimuths are referenced clockwise when viewed from above.",
        ],
    },
]


def draw_centered(c, text, x, y, font="Helvetica", size=7):
    c.setFont(font, size)
    c.drawCentredString(x, y, text)


def draw_callout(c, item_id, x, y, target_x, target_y):
    c.setStrokeColor(BLUE)
    c.setFillColor(colors.white)
    c.setLineWidth(0.8)
    c.line(x, y, target_x, target_y)
    c.circle(x, y, 9, stroke=1, fill=1)
    c.setFillColor(INK)
    draw_centered(c, str(item_id), x, y - 2.4, "Helvetica-Bold", 6.5)


def draw_edge_balloons(c, entries, side, bubble_x=None, bubble_y=None, target_x=None, target_y=None):
    for item_id, point_x, point_y in entries:
        x = point_x if bubble_x is None else bubble_x
        y = point_y if bubble_y is None else bubble_y
        end_x = point_x if target_x is None else target_x
        end_y = point_y if target_y is None else target_y
        c.setStrokeColor(BLUE)
        c.setFillColor(colors.white)
        c.setLineWidth(0.65)
        if side == "top":
            c.line(x, y - 9, end_x, end_y)
        elif side == "bottom":
            c.line(x, y + 9, end_x, end_y)
        elif side == "left":
            c.line(x + 9, y, end_x, end_y)
        else:
            c.line(x - 9, y, end_x, end_y)
        c.circle(x, y, 9, stroke=1, fill=1)
        c.setFillColor(INK)
        draw_centered(c, str(item_id), x, y - 2.4, "Helvetica-Bold", 6.5)


def draw_pump(c):
    c.setStrokeColor(INK)
    c.setLineWidth(1.25)
    # Welded channel base with end plates, grout gap and anchor locations.
    c.line(105, 580, 585, 580)
    c.line(105, 558, 585, 558)
    c.line(105, 580, 105, 558)
    c.line(585, 580, 585, 558)
    for x in (129, 275, 345, 558):
        c.line(x, 580, x, 558)
    c.setLineWidth(0.7)
    for x in (129, 558):
        c.circle(x, 552, 3, stroke=1, fill=0)
        c.line(x, 552, x, 558)
    # Motor frame, cooling fins, fan cover and terminal box.
    c.roundRect(370, 590, 175, 58, 10, stroke=1, fill=0)
    c.roundRect(525, 598, 20, 42, 7, stroke=1, fill=0)
    c.rect(418, 648, 52, 13, stroke=1, fill=0)
    c.rect(430, 661, 28, 8, stroke=1, fill=0)
    for x in range(392, 522, 10):
        c.line(x, 593, x, 645)
    c.line(385, 592, 385, 646)
    c.line(530, 602, 530, 636)
    c.setLineWidth(0.55)
    for y in (598, 604, 610, 616, 622, 628, 634, 640):
        c.line(530, y, 538, y)
    draw_centered(c, "7.5 kW MOTOR", 454, 615, "Helvetica-Bold", 7)
    # Flexible coupling and removable guard.
    c.line(342, 616, 370, 616)
    c.line(342, 621, 370, 621)
    c.roundRect(340, 600, 31, 33, 3, stroke=1, fill=0)
    c.setDash(2, 2)
    c.line(345, 602, 366, 630)
    c.line(345, 630, 366, 602)
    c.setDash()
    # Pump volute casing, bearing bracket and feet.
    c.setFillColor(PALE_BLUE)
    c.circle(300, 616, 38, stroke=1, fill=1)
    c.circle(300, 616, 28, stroke=1, fill=0)
    c.circle(300, 616, 10, stroke=1, fill=0)
    c.setFillColor(colors.white)
    c.rect(284, 578, 12, 12, stroke=1, fill=0)
    c.rect(318, 578, 12, 12, stroke=1, fill=0)
    c.line(338, 616, 342, 616)
    draw_centered(c, "PUMP", 300, 613, "Helvetica-Bold", 6.5)
    # Suction line with eccentric reducer, isolation valve and pressure tap.
    c.setLineWidth(1.25)
    c.line(105, 616, 262, 616)
    c.line(105, 605, 105, 627)
    c.line(112, 605, 112, 627)
    c.line(235, 610, 235, 622)
    c.line(262, 610, 262, 622)
    c.line(235, 610, 249, 616)
    c.line(249, 616, 235, 622)
    c.line(262, 610, 249, 616)
    c.line(249, 616, 262, 622)
    c.line(210, 616, 210, 642)
    c.circle(210, 649, 7, stroke=1, fill=0)
    c.line(210, 642, 210, 649)
    draw_centered(c, "PI", 210, 647, "Helvetica-Bold", 5)
    draw_centered(c, "SUCTION", 153, 630, size=5.5)
    # Discharge elbow, check/isolation valves and flanged spool.
    c.line(300, 654, 300, 680)
    c.line(300, 680, 550, 680)
    c.line(550, 680, 550, 648)
    c.line(300, 654, 314, 654)
    c.line(300, 661, 314, 661)
    for x in (385, 455):
        c.line(x, 674, x, 686)
        c.line(x, 674, x + 10, 680)
        c.line(x + 10, 680, x, 686)
        c.line(x + 10, 674, x + 10, 686)
    c.line(550, 648, 550, 642)
    c.line(544, 642, 556, 642)
    c.line(544, 639, 556, 639)
    draw_centered(c, "DISCHARGE", 482, 691, size=5.5)
    # Hold-downs and package feet.
    for x in (145, 260, 350, 535):
        c.line(x, 558, x, 548)
        c.circle(x, 545, 2.5, stroke=1, fill=0)
    top_ids = (2, 7, 8, 3, 4, 6, 12, 14, 17)
    bottom_ids = (1, 5, 9, 10, 11, 13, 15, 16)
    top_x = (130, 190, 250, 310, 370, 430, 490, 550, 610)
    bottom_x = (145, 205, 265, 325, 385, 445, 505, 565)
    for item_id, x in zip(top_ids, top_x):
        draw_edge_balloons(c, [(item_id, x, 670)], "top", bubble_y=700)
    for item_id, x in zip(bottom_ids, bottom_x):
        draw_edge_balloons(c, [(item_id, x, 558)], "bottom", bubble_y=540)
    c.setStrokeColor(GREY)
    c.setDash(3, 2)
    c.line(105, 525, 585, 525)
    c.setDash()
    draw_centered(c, "1800 BASEPLATE", 345, 512, size=6)
    draw_centered(c, "SIDE ELEVATION  |  NOT TO SCALE", 345, 490, "Helvetica-Bold", 7)
    c.setStrokeColor(INK)
    c.setFillColor(colors.white)
    c.rect(700, 562, 260, 74, stroke=1, fill=1)
    c.circle(830, 599, 28, stroke=1, fill=0)
    c.circle(830, 599, 9, stroke=1, fill=0)
    c.line(700, 599, 802, 599)
    c.line(858, 599, 960, 599)
    c.line(713, 562, 713, 636)
    c.line(947, 562, 947, 636)
    draw_centered(c, "PUMP END VIEW", 830, 540, "Helvetica-Bold", 7)
    draw_centered(c, "DN50 SUCTION", 745, 608, size=6)
    draw_centered(c, "DN32 DISCHARGE", 915, 608, size=6)


def draw_exchanger(c):
    c.setStrokeColor(INK)
    c.setFillColor(colors.white)
    c.setLineWidth(1.3)

    # Sectioned shell, stationary channel and packed floating head.
    c.line(248, 620, 920, 620)
    c.line(248, 500, 920, 500)
    c.line(248, 500, 248, 620)
    c.line(920, 500, 920, 620)
    c.line(235, 500, 235, 620)
    c.line(245, 500, 245, 620)
    c.line(925, 500, 925, 620)
    c.line(935, 500, 935, 620)
    c.line(130, 510, 235, 510)
    c.line(130, 610, 235, 610)
    c.line(130, 510, 130, 610)
    c.arc(78, 510, 132, 610, 90, 180)
    c.line(105, 510, 130, 510)
    c.line(105, 610, 130, 610)

    # Tube bundle and segmental baffles in longitudinal section.
    c.setLineWidth(0.48)
    for y in range(516, 605, 5):
        c.line(250, y, 918, y)
    c.setLineWidth(0.75)
    for index, x in enumerate((320, 400, 480, 560, 640, 720, 800)):
        if index % 2:
            c.line(x, 500, x, 578)
        else:
            c.line(x, 542, x, 620)
    c.line(250, 558, 918, 558)
    c.setStrokeColor(GREY)
    c.setDash(5, 3)
    c.line(70, 558, 1100, 558)
    c.setDash()
    c.setStrokeColor(INK)

    # Two-pass stationary channel with its removable cover and partition.
    c.line(130, 560, 235, 560)
    c.line(183, 510, 183, 560)
    c.line(183, 560, 183, 610)
    c.line(183, 510, 130, 510)
    c.line(183, 610, 130, 610)
    c.line(108, 510, 108, 610)
    c.line(103, 510, 103, 610)
    for y in (520, 535, 550, 565, 580, 595):
        c.circle(240, y, 2, stroke=1, fill=0)

    # Packed floating-head assembly, showing seal stack and removable outer cover.
    c.setLineWidth(1.0)
    c.arc(918, 508, 986, 612, -90, 180)
    c.line(952, 510, 952, 606)
    c.line(958, 510, 958, 606)
    c.line(964, 510, 964, 606)
    c.line(970, 510, 970, 606)
    c.line(978, 504, 978, 614)
    c.arc(974, 498, 1058, 620, -90, 180)
    c.line(1016, 505, 1016, 613)
    c.line(1024, 505, 1024, 613)
    c.setLineWidth(0.65)
    for y in (516, 528, 540, 552, 564, 576, 588, 600):
        c.circle(930, y, 2, stroke=1, fill=0)

    # Expansion joint is shown as a short corrugated shell section.
    c.setLineWidth(1.1)
    c.line(575, 620, 580, 612)
    for x in (580, 589, 598, 607, 616, 625):
        c.line(x, 612, x + 4, 620)
        c.line(x, 500, x + 4, 508)
        c.line(x + 4, 620, x + 9, 612)
        c.line(x + 4, 500, x + 9, 508)
    c.line(634, 620, 639, 612)
    c.line(634, 500, 639, 508)

    # Shell-side nozzles, vent/drain/instrument connections and lifting lug.
    for x, y, upward in ((365, 620, True), (500, 620, True), (740, 620, True),
                         (450, 500, False), (690, 500, False), (840, 620, True)):
        if upward:
            c.rect(x - 6, y, 12, 23, stroke=1, fill=0)
            c.line(x - 10, y + 20, x + 10, y + 20)
            c.line(x - 10, y + 24, x + 10, y + 24)
        else:
            c.rect(x - 6, y - 23, 12, 23, stroke=1, fill=0)
            c.line(x - 10, y - 24, x + 10, y - 24)
            c.line(x - 10, y - 20, x + 10, y - 20)
    c.rect(675, 620, 28, 15, stroke=1, fill=0)
    c.circle(689, 628, 4, stroke=1, fill=0)
    c.rect(136, 545, 18, 26, stroke=1, fill=0)
    c.line(145, 545, 145, 532)
    c.line(139, 532, 151, 532)

    # Floating-head pass partition, tie rods, impingement plate and saddle supports.
    c.setLineWidth(0.75)
    c.line(880, 558, 918, 558)
    c.line(880, 558, 880, 540)
    for y in (521, 535, 579, 593):
        c.line(260, y, 916, y)
    c.setFillColor(PALE_BLUE)
    c.rect(250, 568, 14, 20, stroke=1, fill=1)
    c.setFillColor(colors.white)
    for x in (385, 755):
        c.line(x - 34, 500, x - 20, 468)
        c.line(x + 34, 500, x + 20, 468)
        c.line(x - 20, 468, x + 20, 468)
        c.rect(x - 42, 463, 84, 5, stroke=1, fill=0)
        c.line(x - 36, 457, x + 36, 457)
        c.line(x - 36, 457, x - 36, 463)
        c.line(x + 36, 457, x + 36, 463)

    # Callouts follow the numbered component schedule on sheet 2.
    top_ids = (4, 5, 2, 1, 3, 6, 10, 30, 7, 33, 28, 8, 12, 29, 31, 14, 37, 35, 11)
    top_targets = (108, 145, 160, 185, 235, 248, 250, 257, 350, 365, 425, 490, 500, 540,
                   560, 605, 689, 740, 920)
    bottom_ids = (19, 36, 34, 15, 32, 21, 9, 38, 13, 27, 18, 22, 17, 24, 23, 25, 16, 26, 20)
    bottom_targets = (145, 385, 450, 880, 880, 910, 920, 930, 935, 952, 958, 964, 970, 974,
                      978, 986, 1016, 1024, 1035)
    top_x = [92 + index * 56 for index in range(len(top_ids))]
    bottom_x = [92 + index * 56 for index in range(len(bottom_ids))]
    for item_id, x, target in zip(top_ids, top_x, top_targets):
        draw_edge_balloons(c, [(item_id, x, 650)], "top", bubble_y=690, target_x=target, target_y=620)
    for item_id, x, target in zip(bottom_ids, bottom_x, bottom_targets):
        draw_edge_balloons(c, [(item_id, x, 500)], "bottom", bubble_y=420, target_x=target, target_y=500)

    c.setFillColor(INK)
    draw_centered(c, "STATIONARY CHANNEL", 155, 480, "Helvetica-Bold", 6.5)
    draw_centered(c, "SHELL / TUBE BUNDLE SECTION", 585, 445, "Helvetica-Bold", 7)
    draw_centered(c, "PACKED FLOATING HEAD", 990, 480, "Helvetica-Bold", 6.5)
    c.setStrokeColor(GREY)
    c.setDash(3, 2)
    c.line(125, 408, 1055, 408)
    c.setDash()
    draw_centered(c, "APPROX. 3200 FACE-TO-FACE  |  TWO-PASS TUBE SIDE", 585, 393, size=6)


def draw_receiver(c):
    c.setStrokeColor(INK)
    c.setFillColor(PALE_BLUE)
    c.setLineWidth(1.4)
    vessel = c.beginPath()
    vessel.moveTo(330, 660)
    vessel.curveTo(330, 682, 357, 695, 390, 695)
    vessel.curveTo(423, 695, 450, 682, 450, 660)
    vessel.lineTo(450, 555)
    vessel.curveTo(450, 533, 423, 520, 390, 520)
    vessel.curveTo(357, 520, 330, 533, 330, 555)
    vessel.close()
    c.drawPath(vessel, stroke=1, fill=1)
    c.setFillColor(colors.white)
    # Top instrument nozzle and flanged relief-valve branch.
    c.rect(383, 695, 14, 14, stroke=1, fill=0)
    c.line(379, 709, 401, 709)
    c.line(379, 712, 401, 712)
    c.line(420, 695, 420, 704)
    c.line(414, 704, 426, 704)
    c.line(414, 707, 426, 707)
    c.circle(420, 714, 7, stroke=1, fill=0)
    c.line(420, 707, 420, 714)
    # Reinforced side nozzles, flanges and manway neck.
    for x, y, direction in ((330, 625, -1), (450, 590, 1)):
        c.rect(x - 4, y - 14, 8, 28, stroke=1, fill=0)
        c.line(x, y - 10, x + direction * 24, y - 10)
        c.line(x, y + 10, x + direction * 24, y + 10)
        c.line(x + direction * 24, y - 14, x + direction * 24, y + 14)
        c.line(x + direction * 28, y - 14, x + direction * 28, y + 14)
        c.line(x + direction * 24, y - 14, x + direction * 28, y - 14)
        c.line(x + direction * 24, y + 14, x + direction * 28, y + 14)
    c.rect(326, 595, 4, 30, stroke=1, fill=0)
    c.circle(330, 610, 19, stroke=1, fill=0)
    for angle in range(0, 360, 45):
        bx = 330 + 16 * math.cos(math.radians(angle))
        by = 610 + 16 * math.sin(math.radians(angle))
        c.circle(bx, by, 1.1, stroke=1, fill=0)
    # Skirt support, base ring, anchor bolts and vented drainage gap.
    c.line(350, 520, 350, 488)
    c.line(430, 520, 430, 488)
    c.line(350, 488, 430, 488)
    c.line(342, 488, 438, 488)
    c.line(342, 484, 438, 484)
    for x in (350, 430):
        c.line(x, 484, x, 475)
        c.circle(x, 473, 2, stroke=1, fill=0)
    # Shell seams and a vertical centerline.
    c.setStrokeColor(GREY)
    c.setDash(2, 2)
    c.line(390, 505, 390, 720)
    c.setDash()
    c.setStrokeColor(INK)
    c.setLineWidth(0.6)
    c.line(331, 660, 449, 660)
    c.line(331, 555, 449, 555)
    c.line(335, 625, 445, 625)
    c.line(335, 590, 445, 590)
    c.setFillColor(INK)
    draw_centered(c, "V-0805", 390, 606, "Helvetica-Bold", 8)
    draw_centered(c, "ELEVATION  |  O/A HEIGHT 3200", 390, 455, "Helvetica-Bold", 6.5)
    # Overall height dimension and extension lines.
    c.setFillColor(colors.white)
    c.setStrokeColor(GREY)
    c.setDash(3, 2)
    c.line(310, 484, 310, 695)
    c.setDash()
    c.line(305, 484, 315, 484)
    c.line(305, 695, 315, 695)
    c.line(310, 484, 320, 494)
    c.line(310, 695, 320, 685)
    left_callouts = (
        (1, 300, 690), (2, 300, 660), (6, 300, 630), (13, 300, 600),
        (3, 300, 570), (4, 300, 540), (5, 300, 510), (16, 300, 480),
    )
    right_callouts = (
        (8, 480, 690), (9, 480, 660), (10, 480, 630), (7, 480, 600),
        (11, 480, 570), (12, 480, 540), (14, 480, 510), (15, 480, 480),
    )
    draw_edge_balloons(c, left_callouts, "left", bubble_x=270)
    draw_edge_balloons(c, right_callouts, "right", bubble_x=510)
    c.setStrokeColor(INK)
    c.setFillColor(colors.white)
    cx, cy = 885, 522
    c.circle(cx, cy, 43, stroke=1, fill=0)
    c.circle(cx, cy, 37, stroke=1, fill=0)
    c.setStrokeColor(GREY)
    c.setDash(3, 2)
    c.line(cx - 51, cy, cx + 51, cy)
    c.line(cx, cy - 51, cx, cy + 51)
    c.setDash()
    c.setStrokeColor(INK)
    nozzle_angles = (90, 30, -30, -90, -150, 150)
    for index, angle in enumerate(nozzle_angles, start=1):
        radians = math.radians(angle)
        inner_x = cx + 37 * math.cos(radians)
        inner_y = cy + 37 * math.sin(radians)
        outer_x = cx + 49 * math.cos(radians)
        outer_y = cy + 49 * math.sin(radians)
        c.line(inner_x, inner_y, outer_x, outer_y)
        c.circle(outer_x, outer_y, 4, stroke=1, fill=0)
        label_x = cx + 59 * math.cos(radians)
        label_y = cy + 59 * math.sin(radians) - 2
        draw_centered(c, f"N{index}", label_x, label_y, "Helvetica-Bold", 5.5)
    c.setFillColor(INK)
    draw_centered(c, "PLAN / NOZZLE ORIENTATION", cx, 455, "Helvetica-Bold", 7)


def draw_views(c, sample):
    c.setStrokeColor(GRID)
    c.setLineWidth(0.5)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(GREY)
    c.setFont("Helvetica", 6)
    if sample["kind"] == "pump":
        c.rect(55, 470, 650, 280, stroke=1, fill=0)
        c.rect(725, 470, 410, 280, stroke=1, fill=0)
        c.setFillColor(INK)
        c.drawString(66, 735, "MAIN ARRANGEMENT")
        c.drawString(736, 735, "DETAIL / CONNECTION VIEW")
        c.setFillColor(GREY)
        c.drawString(66, 720, f"EQUIPMENT: {sample['equipment']}")
        draw_pump(c)
    elif sample["kind"] == "exchanger":
        c.rect(55, 330, 1080, 420, stroke=1, fill=0)
        c.setFillColor(INK)
        c.drawString(66, 735, "LONGITUDINAL SECTION / PACKED FLOATING HEAD")
        c.setFillColor(GREY)
        c.drawString(66, 720, f"EQUIPMENT: {sample['equipment']}")
        draw_exchanger(c)
    else:
        c.rect(55, 470, 650, 280, stroke=1, fill=0)
        c.rect(725, 470, 410, 280, stroke=1, fill=0)
        c.setFillColor(INK)
        c.drawString(66, 735, "MAIN ARRANGEMENT")
        c.drawString(736, 735, "DETAIL / CONNECTION VIEW")
        c.setFillColor(GREY)
        c.drawString(66, 720, f"EQUIPMENT: {sample['equipment']}")
        draw_receiver(c)


def draw_parts_list(c, items, top=438, row_h=12):
    x0, x1 = 55, 1135
    columns = [
        ("ITEM", 64),
        ("DESCRIPTION", 103),
        ("MATERIAL", 445),
        ("SIZE", 605),
        ("QTY", 802),
        ("PART NO", 875),
    ]
    bottom = top - row_h * (len(items) + 1)
    c.setStrokeColor(INK)
    c.setLineWidth(0.7)
    c.rect(x0, bottom, x1 - x0, top - bottom, stroke=1, fill=0)
    for index in range(len(items) + 2):
        y = top - row_h * index
        c.setStrokeColor(GRID if index else INK)
        c.setLineWidth(0.5)
        c.line(x0, y, x1, y)
    c.setFillColor(PALE_BLUE)
    c.rect(x0, top - row_h, x1 - x0, row_h, stroke=0, fill=1)
    c.setFillColor(INK)
    for header, x in columns:
        c.setFont("Helvetica-Bold", 6.2)
        c.drawString(x, top - 8.3, header)
    c.setFont("Helvetica", 6.0)
    col_x = [x for _, x in columns]
    for row_index, part in enumerate(items, start=1):
        y = top - row_h * (row_index + 0.72)
        vals = (
            part["item"],
            part["description"],
            part["material"],
            part["size"],
            str(part["qty"]),
            part["partno"],
        )
        for index, (value, x) in enumerate(zip(vals, col_x)):
            max_width = (col_x[index + 1] - x - 7) if index + 1 < len(col_x) else x1 - x - 7
            text = value
            while text and stringWidth(text, "Helvetica", 6.0) > max_width:
                text = text[:-1]
            if text != value:
                text = text.rstrip() + "..."
            c.drawString(x, y, text)
    c.setFillColor(GREY)
    c.setFont("Helvetica-Oblique", 5.8)
    c.drawString(x0, bottom - 11, "BILL OF MATERIALS  |  QUANTITIES ARE PER COMPLETE ASSEMBLY")


def draw_title_block(c, sample, sheet="1 OF 1"):
    x, y, width, height = 55, 38, 1080, 82
    c.setStrokeColor(INK)
    c.setLineWidth(0.9)
    c.rect(x, y, width, height, stroke=1, fill=0)
    c.line(x, y + 28, x + width, y + 28)
    c.line(x + 700, y, x + 700, y + height)
    c.line(x + 880, y, x + 880, y + height)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(x + 10, y + 64, "NORTHBRIDGE THERMAL SYSTEMS")
    c.setFillColor(GREY)
    c.setFont("Helvetica", 5.5)
    c.drawString(x + 10, y + 53, "ENGINEERING PROTOTYPE  |  FICTIONAL EQUIPMENT AND DATA")
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x + 10, y + 37, sample["title"])
    c.setFont("Helvetica", 6)
    c.drawString(x + 10, y + 17, sample["subtitle"])
    labels = [
        ("DRAWING NUMBER", sample["drawing_no"]),
        ("REVISION", sample["revision"]),
        ("SCALE", "NTS"),
        ("SHEET", sheet),
    ]
    for i, (label, value) in enumerate(labels):
        col = i % 2
        row = i // 2
        cx = x + 713 + col * 88
        cy = y + 59 - row * 34
        c.setFillColor(GREY)
        c.setFont("Helvetica", 5)
        c.drawString(cx, cy, label)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(cx, cy - 11, value)
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 6)
    c.drawString(x + 894, y + 56, "PROTOTYPE SAMPLE")
    c.setFont("Helvetica", 5.5)
    c.drawString(x + 894, y + 43, "NOT FOR FABRICATION")
    c.drawString(x + 894, y + 29, "Verify all design data")
    c.drawString(x + 894, y + 17, "before engineering use")


def generate_pdf(sample, items, output_path):
    c = canvas.Canvas(str(output_path), pagesize=(PAGE_W, PAGE_H), pageCompression=1)
    c.setTitle(f"{sample['drawing_no']} - {sample['title']}")
    c.setAuthor("Northbridge Thermal Systems - Fictional Prototype Sample")
    c.setSubject("Synthetic engineering drawing for BOM reviewer prototype demonstration")
    c.setStrokeColor(GRID)
    c.setLineWidth(0.5)
    c.rect(28, 28, PAGE_W - 56, PAGE_H - 56, stroke=1, fill=0)

    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(55, PAGE_H - 53, sample["title"])
    c.setFillColor(GREY)
    c.setFont("Helvetica", 7)
    c.drawString(55, PAGE_H - 68, sample["equipment"])
    c.drawRightString(PAGE_W - 56, PAGE_H - 53, f"DRAWING: {sample['drawing_no']}    REV: {sample['revision']}")
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 6)
    c.drawRightString(PAGE_W - 56, PAGE_H - 68, "FICTIONAL PROTOTYPE - NOT FOR FABRICATION")
    draw_views(c, sample)
    if sample["kind"] != "exchanger":
        draw_parts_list(c, items)
    notes_y = 148
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 6.2)
    c.drawString(55, notes_y + 11, "GENERAL NOTES")
    c.setFillColor(GREY)
    c.setFont("Helvetica", 5.8)
    for i, note in enumerate(sample["notes"], start=1):
        c.drawString(55, notes_y - i * 9, f"{i}. {note}")
    if sample["kind"] == "exchanger":
        c.setFillColor(BLUE)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(55, 310, "NUMBERED COMPONENTS AND FULL BOM SCHEDULE: SHEET 2 OF 2")
        draw_title_block(c, sample, "1 OF 2")
        c.showPage()
        draw_exchanger_schedule_page(c, sample, items)
    else:
        draw_title_block(c, sample)
    c.showPage()
    c.save()


def draw_exchanger_schedule_page(c, sample, items):
    c.setStrokeColor(GRID)
    c.setLineWidth(0.5)
    c.rect(28, 28, PAGE_W - 56, PAGE_H - 56, stroke=1, fill=0)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(55, PAGE_H - 53, "NUMBERED COMPONENT SCHEDULE AND BILL OF MATERIALS")
    c.setFillColor(GREY)
    c.setFont("Helvetica", 7)
    c.drawString(55, PAGE_H - 68, f"{sample['equipment']}  |  ITEM NUMBERS MATCH SHEET 1 CALLOUTS")
    c.drawRightString(PAGE_W - 56, PAGE_H - 53, f"DRAWING: {sample['drawing_no']}    REV: {sample['revision']}")
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 6)
    c.drawRightString(PAGE_W - 56, PAGE_H - 68, "FICTIONAL PROTOTYPE - NOT FOR FABRICATION")
    draw_parts_list(c, items, top=PAGE_H - 125, row_h=13)
    c.setFillColor(GREY)
    c.setFont("Helvetica", 5.8)
    c.drawString(55, 148, "Drawing schedule is included for extraction testing; verify all quantities and specifications before engineering use.")
    draw_title_block(c, sample, "2 OF 2")


def adjusted_bom(sample):
    rows = []
    for part in sample["items"]:
        if part["item"] in sample["missing_bom"]:
            continue
        row = dict(part)
        row.update(sample["bom_changes"].get(part["item"], {}))
        if part["item"] not in sample["bom_changes"] or row != part:
            rows.append(row)
    rows.extend(dict(extra) for extra in sample["extra_bom"])
    return rows


def generate_xlsx(sample, items, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    ws["A1"] = "NORTHBRIDGE THERMAL SYSTEMS  |  EQUIPMENT MATERIALS LIST"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="183047")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:G2")
    ws["A2"] = sample["title"]
    ws["A2"].font = Font(name="Arial", size=11, bold=True, color="1F587C")
    ws.merge_cells("A3:C3")
    ws["A3"] = f"Equipment: {sample['equipment']}"
    ws.merge_cells("D3:G3")
    ws["D3"] = f"Drawing: {sample['drawing_no']}"
    ws.merge_cells("A4:C4")
    ws["A4"] = "FICTIONAL PROTOTYPE SAMPLE - NOT FOR PROCUREMENT"
    ws["A4"].font = Font(name="Arial", size=9, bold=True, color="C2413B")
    ws.merge_cells("D4:E4")
    ws["D4"] = f"Revision: {sample['revision']}"
    ws.merge_cells("F4:G4")
    ws["F4"] = "Units: mm unless noted"
    ws.merge_cells("A5:G5")
    ws["A5"] = "Demonstration data only. Confirm specifications and quantities before engineering use."
    ws["A5"].font = Font(name="Arial", size=9, italic=True, color="657786")

    headers = ["Item No", "Description", "Material", "Size / Specification", "Qty", "Part No", "Remarks"]
    header_row = 7
    for column, value in enumerate(headers, start=1):
        cell = ws.cell(header_row, column, value)
        cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F587C")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color="183047"))
    ws.row_dimensions[header_row].height = 30

    bom_rows = adjusted_bom(sample)
    for row_number, part in enumerate(bom_rows, start=header_row + 1):
        values = [
            part["item"],
            part["description"],
            part["material"],
            part["size"],
            part["qty"],
            part["partno"],
            part["remarks"] or ("Spare item" if part["item"] == "SP-01" else ""),
        ]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row_number, column, value)
            cell.font = Font(name="Arial", size=9, color="183047")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="d8e0e7"))
            if row_number % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F3F7FA")
        ws.row_dimensions[row_number].height = 22

    widths = [12, 41, 25, 26, 9, 20, 22]
    for column, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:G{header_row + len(bom_rows)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"1:{header_row}"
    ws.sheet_properties.outlinePr.summaryBelow = True

    revision = wb.create_sheet("Revision Log")
    revision.sheet_view.showGridLines = False
    revision.append(["Revision", "Description", "Prepared By", "Date"])
    revision.append([sample["revision"], "Prototype demonstration issue", "Northbridge Engineering", "2026-10-10"])
    for cell in revision[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F587C")
    revision.column_dimensions["A"].width = 14
    revision.column_dimensions["B"].width = 42
    revision.column_dimensions["C"].width = 28
    revision.column_dimensions["D"].width = 18
    wb.save(output_path)


def generate_sample(sample):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    pdf_path = OUTPUT_DIR / f"{sample['stem']}.pdf"
    xlsx_path = OUTPUT_DIR / f"{sample['stem']}_bom.xlsx"
    generate_pdf(sample, sample["items"], pdf_path)
    generate_xlsx(sample, sample["items"], xlsx_path)
    return pdf_path, xlsx_path


def main():
    for sample in SAMPLES:
        pdf_path, xlsx_path = generate_sample(sample)
        print(f"Created {pdf_path.name} and {xlsx_path.name}")


if __name__ == "__main__":
    main()
