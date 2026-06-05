"""Shared constants — sheet names, column letters, dropdown options, headers."""
from __future__ import annotations

SHEET_ALL_ORDERS = "All Orders"
SHEET_ITEM_LIST = "item list (Web CSV)"
SHEET_INVENTORY = "Inventory"
SHEET_SHOPPING_HISTORY = "Shopping History"

SHEET_JETRO_SOURCE = "Jetro source"
SHEET_PO = "PO"
SHEET_WAREHOUSE_SHORT = "Warehouse short"
SHEET_DRIVER_SETUP = "Driver Setup"

SOURCE_SHEETS = (SHEET_INVENTORY, SHEET_ITEM_LIST, SHEET_SHOPPING_HISTORY)
ROUTABLE_SHEETS = (SHEET_ALL_ORDERS, SHEET_JETRO_SOURCE, SHEET_PO)

# Bin values valid on All Orders
BIN_COOLER = "COOLER"
BIN_COOLER_PD = "COOLER-PD"
BIN_DRY = "DRY"
BIN_FREEZER = "FREEZER"
BIN_PICK_UP = "PICK UP"
ALL_ORDERS_BINS = (BIN_COOLER, BIN_COOLER_PD, BIN_DRY, BIN_FREEZER, BIN_PICK_UP)

# WH Pickup bin sort priority — anything else after FREEZER, then alphabetically.
WH_PICKUP_BIN_ORDER = {
    BIN_COOLER: 0,
    BIN_COOLER_PD: 1,
    BIN_DRY: 2,
    BIN_FREEZER: 3,
}

# Compile column mapping (Part 1 §3.1).
# Each entry: (source sheet, source header, target column letter, target header).
COMPILE_MAP = [
    (SHEET_ITEM_LIST,        "Category Name",          "T",  "CATEGORY NAME"),
    (SHEET_ITEM_LIST,        "Current Cost Price",     "U",  "CURRENT COST PRICE"),
    (SHEET_ITEM_LIST,        "Current Selling Price",  "V",  "CURRENT SELLING PRICE"),
    (SHEET_ITEM_LIST,        "Unit",                   "AF", "UNIT"),
    (SHEET_SHOPPING_HISTORY, "Product",                "W",  "Product"),
    (SHEET_SHOPPING_HISTORY, "Bin",                    "X",  "New bin"),
    (SHEET_SHOPPING_HISTORY, "Unit Price",             "Y",  "unit price"),
    (SHEET_SHOPPING_HISTORY, "Case Price",             "Z",  "case price"),
    (SHEET_INVENTORY,        "Quantity On Hand",       "AA", "Quantity On Hand"),
    (SHEET_INVENTORY,        "Cost",                   "AB", "Cost"),
    (SHEET_INVENTORY,        "Case Avg Weight",        "AC", "CASE AVG WEIGHT"),
    (SHEET_INVENTORY,        "BIN (Internal)",         "AD", "BIN(Internal)"),
]

# Source code columns used to match against All Orders.Code.
CODE_COLUMN_BY_SHEET = {
    SHEET_ALL_ORDERS:        "Code",
    SHEET_ITEM_LIST:         "SKU CODE",
    SHEET_INVENTORY:         "Item",
    SHEET_SHOPPING_HISTORY:  "Item",
}

# Number format for all price and weight columns.
NUMBER_FORMAT_PRECISION = "0.####"
PRECISION_COLUMNS = {
    "CURRENT COST PRICE",
    "CURRENT SELLING PRICE",
    "unit price",
    "case price",
    "Cost",
    "CASE AVG WEIGHT",
    "Price",
}

# Default Jetro driver pull sequence (fallback when no Driver Setup is given).
DEFAULT_JETRO_DRIVER_ORDER = (
    "Glen", "Abraham", "Ramon", "Gilberto", "Jose", "Adrian", "Luis", "Z",
)

# UNIT classifications.
UNIT_LBS = "LBS"
UNIT_CASE = "CASE"
UNIT_UNIT = "Unit"

# Warehouse short — Update vendor routing keywords.
WH_ROUTING_JETRO = "Jetro"
WH_ROUTING_WH = "WH"
