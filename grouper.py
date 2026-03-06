"""
grouper.py — Cabinet & Room Grouping Logic
Takes a classified BOM dataframe and builds a nested structure:

  { room_name: { cabinet_name: [ component_dict, ... ] } }

This structure drives the layout and drawing engines.
"""

import json
import math
import os

RULES_PATH = os.path.join(os.path.dirname(__file__), 'rules.json')

# ── IO Cabinet sizing rules (from architecture engine R1 / R2 / R3) ───────────
# R1: 1 CIOC per 8 CHARM baseplates (one tower)
# R2: Max 3 towers per IO cabinet
# R3: IO cabinet count = ceiling(total_baseplates / 24)
CHARMS_PER_TOWER  = 8
TOWERS_PER_IO_CAB = 3
CHARMS_PER_IO_CAB = CHARMS_PER_TOWER * TOWERS_PER_IO_CAB   # = 24


# ── Cabinet display names ──────────────────────────────────────────────────────
CABINET_LABELS = {
    'SERVER_CABINET': 'Server Cabinet',
    'IO_CABINET':     'I/O Cabinet',
    'OPERATOR_DESK':  'Operator Desk',
    'SOFTWARE':       'SOFTWARE_STRIP',   # rendered as a special strip, not a cabinet
    'META':           None,               # CABINETs themselves — skip from diagram items
}

# ── Room mapping ───────────────────────────────────────────────────────────────
def _canonical_room(area: str) -> str:
    a = str(area).strip().upper()
    if 'OPERATOR' in a:
        return 'OPERATOR ROOM'
    return 'PDC ROOM'


def _cabinet_for_class(diagram_class: str, rules: dict) -> str:
    """Get the cabinet_group for a diagram_class from rules.json."""
    if diagram_class in rules:
        return rules[diagram_class].get('cabinet_group', 'SERVER_CABINET')
    return 'SERVER_CABINET'


def _count_charm_baseplates(items: list) -> int:
    """
    Calculate total CHARM baseplates needed from IO items.

    Two cases found in BOMs:
      1. CHARM_BASE rows   — qty IS the baseplate count (e.g. qty=8 → 8 baseplates)
      2. CHARM_AI/AO/DI/DO — qty is individual card count; 8 cards share 1 baseplate
                              (e.g. 96 AI cards → ceil(96/8) = 12 baseplates)

    If CHARM_BASE rows exist we trust them as ground truth.
    Otherwise we derive from individual card counts.
    """
    base_rows = [it for it in items if it.get('diagram_class') == 'CHARM_BASE']
    card_rows = [it for it in items
                 if it.get('diagram_class', '').startswith('CHARM_')
                 and it.get('diagram_class') != 'CHARM_BASE']

    if base_rows:
        # BOM explicitly lists baseplates — use directly
        return sum(it.get('qty', 1) for it in base_rows)

    if card_rows:
        # BOM lists individual IO cards — derive baseplates (8 cards per baseplate)
        total_cards = sum(it.get('qty', 1) for it in card_rows)
        return math.ceil(total_cards / CHARMS_PER_TOWER)

    return 0


# Keep old name as alias so nothing else breaks
def _count_charm_units(items: list) -> int:
    return _count_charm_baseplates(items)


def _auto_insert_ciocs(io_items: list) -> list:
    """
    R1: If CIOCs are absent (or fewer than needed) auto-insert them.
    Rule: 1 CIOC per 8 CHARM baseplates (1 tower).
    Uses _count_charm_baseplates so card-based BOMs are handled correctly.
    Injected items are marked confidence='AUTO' so the UI can show them.
    """
    existing_ciocs = [it for it in io_items if it['diagram_class'] == 'CIOC']
    baseplate_total = _count_charm_baseplates(io_items)
    expected        = math.ceil(baseplate_total / CHARMS_PER_TOWER) if baseplate_total > 0 else 0

    shortfall = expected - len(existing_ciocs)
    for i in range(shortfall):
        io_items.append({
            'description':  f'CIOC (Auto #{len(existing_ciocs) + i + 1})',
            'qty':          1,
            'diagram_class':'CIOC',
            'part_number':  '',
            'label':        'CIOC',
            'color_hex':    '2E5FA3',
            'text_color':   'FFFFFF',
            'confidence':   'AUTO',
        })
    return io_items


def _expand_redundant_controllers(items: list) -> list:
    """
    R4: A single 'Redundant ...' CONTROLLER row becomes two rows:
    PRI CNTR and SEC CNTR, each qty=1.
    """
    result = []
    for item in items:
        if (item.get('diagram_class') == 'CONTROLLER'
                and 'redundant' in item.get('description', '').lower()):
            for role, tag in [('PRIMARY', 'PRI'), ('SECONDARY', 'SEC')]:
                clone = dict(item)
                clone['description']    = f'{tag} CNTR'
                clone['qty']            = 1
                clone['redundant_role'] = role
                result.append(clone)
        else:
            result.append(item)
    return result


def _split_io_cabinets(io_items: list) -> dict:
    """
    R4: Expand redundant controllers into PRI + SEC.
    R1: Auto-insert CIOCs (1 per 8 CHARMs) if missing from BOM.
    R3: Split across IO cabinets using 24-baseplate rule (not 50/50).
         ceiling(total_charms / 24) cabinets.
         Priority items (CONTROLLER, CIOC, POWER) always go to Cabinet #1.
         CHARM items distributed proportionally across cabinets.
    """
    # R4: Expand redundant controllers
    io_items = _expand_redundant_controllers(io_items)
    # R1: Auto-insert CIOCs
    io_items = _auto_insert_ciocs(io_items)

    charm_total    = _count_charm_baseplates(io_items)
    priority_cls   = {'CONTROLLER', 'CIOC', 'POWER'}
    priority_items = [it for it in io_items if it['diagram_class'] in priority_cls]
    charm_items    = [it for it in io_items if 'CHARM' in it.get('diagram_class', '')]
    other_items    = [it for it in io_items
                      if it['diagram_class'] not in priority_cls
                      and 'CHARM' not in it.get('diagram_class', '')]

    # R3: Number of IO cabinets
    num_cabs = max(1, math.ceil(charm_total / CHARMS_PER_IO_CAB)) if charm_total > 0 else 1

    if num_cabs == 1:
        return {'I/O Cabinet #1': priority_items + charm_items + other_items}

    # Distribute CHARM items across cabinets (budget = 24 per cabinet)
    result     = {}
    charm_pool = list(charm_items)   # mutable copy
    ci         = 0                   # index into charm_pool

    for cab_idx in range(num_cabs):
        cab_name  = f'I/O Cabinet #{cab_idx + 1}'
        cab_items = []

        if cab_idx == 0:
            cab_items.extend(priority_items)

        budget = CHARMS_PER_IO_CAB
        placed = 0

        while ci < len(charm_pool) and placed < budget:
            item       = charm_pool[ci]
            item_qty   = item.get('qty', 1)
            remaining  = budget - placed

            if item_qty <= remaining:
                cab_items.append(item)
                placed += item_qty
                ci     += 1
            else:
                # Split item across cabinet boundary
                part1          = dict(item); part1['qty'] = remaining
                part2          = dict(item); part2['qty'] = item_qty - remaining
                cab_items.append(part1)
                charm_pool[ci] = part2
                placed         = budget   # cabinet is full

        result[cab_name] = cab_items

    return result


def group_bom(df, rules_path=RULES_PATH) -> dict:
    """
    Group a classified BOM dataframe into a nested room/cabinet structure.

    Returns:
    {
      'PDC ROOM': {
          'SERVER_CABINET': [
              {'description': ..., 'qty': ..., 'diagram_class': ...,
               'label': ..., 'color_hex': ..., 'text_color': ...}, ...
          ],
          'IO_CABINET_1': [...],
          'IO_CABINET_2': [...],
          'SOFTWARE': [...]          # special — top strip
      },
      'OPERATOR ROOM': {
          'OPERATOR_DESK': [...]
      }
    }
    """
    with open(rules_path, 'r') as f:
        rules = json.load(f)

    # Initialize structure
    structure = {
        'PDC ROOM':      {'SERVER_CABINET': [], 'IO_CABINET': [], 'SOFTWARE': [], 'FOPP_NODES': []},
        'OPERATOR ROOM': {'OPERATOR_DESK': []}
    }

    for _, row in df.iterrows():
        dclass  = str(row.get('diagram_class', 'UNKNOWN')).strip()
        area    = _canonical_room(str(row.get('area', 'PDC ROOM')))
        cabinet = _cabinet_for_class(dclass, rules)

        # META (cabinet enclosures themselves) — skip
        if cabinet == 'META':
            continue

        # Get display info from rules
        rule_info = rules.get(dclass, {})
        item = {
            'description':  str(row.get('description', '')).strip(),
            'qty':          int(row.get('qty', 1)),
            'diagram_class': dclass,
            'part_number':  str(row.get('part_number', '')).strip(),
            'label':        rule_info.get('display_label', dclass),
            'color_hex':    rule_info.get('color_hex', '888888'),
            'text_color':   rule_info.get('text_color', 'FFFFFF'),
            'confidence':   str(row.get('confidence', 'MEDIUM')),
        }

        # Route to correct room & cabinet
        if area == 'OPERATOR ROOM':
            if cabinet == 'OPERATOR_DESK':
                structure['OPERATOR ROOM']['OPERATOR_DESK'].append(item)
            else:
                # Anything in operator room that isn't explicitly OPERATOR_DESK
                # still goes there
                structure['OPERATOR ROOM']['OPERATOR_DESK'].append(item)

        else:  # PDC ROOM
            if cabinet == 'SOFTWARE':
                structure['PDC ROOM']['SOFTWARE'].append(item)
            elif cabinet == 'IO_CABINET':
                structure['PDC ROOM']['IO_CABINET'].append(item)
            # R5 (other team): FOPP is a standalone topology node, not a cabinet component
            elif cabinet == 'FOPP_NODES':
                structure['PDC ROOM']['FOPP_NODES'].append(item)
            else:
                structure['PDC ROOM']['SERVER_CABINET'].append(item)

    # Split IO cabinet items across physical cabinets
    io_items = structure['PDC ROOM'].pop('IO_CABINET')
    io_split  = _split_io_cabinets(io_items)
    structure['PDC ROOM'].update(io_split)

    # Remove empty cabinets (but keep FOPP_NODES even if empty — generator needs it)
    for room in list(structure.keys()):
        for cab in list(structure[room].keys()):
            if not structure[room][cab] and cab != 'FOPP_NODES':
                del structure[room][cab]
        if not structure[room]:
            del structure[room]

    return structure


def summarize(structure: dict):
    """Print a human-readable summary of the grouped structure."""
    for room, cabinets in structure.items():
        print(f"\n📦 {room}")
        for cab, items in cabinets.items():
            print(f"   🗄️  {cab}  ({len(items)} items)")
            for it in items:
                print(f"       • [{it['diagram_class']:12s}] {it['description'][:60]}  qty={it['qty']}")


if __name__ == '__main__':
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from parser     import parse_bom
    from classifier import classify_dataframe

    path = sys.argv[1] if len(sys.argv) > 1 else '../bom_full.csv'
    df   = parse_bom(path)
    df   = classify_dataframe(df)
    grp  = group_bom(df)
    summarize(grp)