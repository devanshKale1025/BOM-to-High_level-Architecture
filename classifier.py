"""
classifier.py — Rule-Based Classification Engine
Assigns a Diagram_Class to every BOM row using 3 levels of matching.

Level 1 — Part number prefix match   (highest confidence)
Level 2 — Exact phrase match
Level 3 — Keyword match              (fallback)
"""

import json
import os
import re

RULES_PATH = os.path.join(os.path.dirname(__file__), 'rules.json')


def load_rules(path=RULES_PATH):
    with open(path, 'r') as f:
        return json.load(f)


def _normalize(text: str) -> str:
    """Lowercase, strip extra spaces, remove punctuation noise."""
    t = str(text).lower().strip()
    t = re.sub(r'[;:,]', ' ', t)
    t = re.sub(r'\s+', ' ', t)
    return t


def classify_row(description: str, part_number: str, category: str, area: str,
                 rules: dict) -> tuple[str, str]:
    """
    Classify one BOM row.

    Returns:
        (diagram_class, confidence)
        confidence: 'HIGH' | 'MEDIUM' | 'LOW' | 'UNKNOWN'
    """
    desc_norm = _normalize(description)
    part_norm = _normalize(part_number)
    cat_norm  = _normalize(category)
    area_norm = _normalize(area)

    # ── Special case: Operator Room workstation must be OPERATOR_WS ──────────
    if 'operator' in area_norm and ('workstation' in desc_norm or 'tower' in desc_norm):
        return 'OPERATOR_WS', 'HIGH'

    # ── Level 1: Part number prefix match ────────────────────────────────────
    for cls, rule in rules.items():
        for prefix in rule.get('part_prefixes', []):
            if part_norm.startswith(prefix.lower()):
                return cls, 'HIGH'

    # ── Level 2: Exact phrase match in description ────────────────────────────
    for cls, rule in rules.items():
        for kw in rule.get('keywords', []):
            if _normalize(kw) == desc_norm:
                return cls, 'HIGH'

    # ── Level 3: Keyword substring match ─────────────────────────────────────
    # Score each class by how many keywords match
    scores = {}
    for cls, rule in rules.items():
        score = 0
        for kw in rule.get('keywords', []):
            kw_norm = _normalize(kw)
            if kw_norm in desc_norm:
                # Longer keyword match = higher score
                score += len(kw_norm.split())
        if score > 0:
            scores[cls] = score

    if scores:
        best = max(scores, key=scores.get)
        confidence = 'MEDIUM' if scores[best] >= 2 else 'LOW'
        return best, confidence

    # ── Category fallback ─────────────────────────────────────────────────────
    cat_map = {
        'cntr':  'CONTROLLER', 'ws':   'WORKSTATION', 'swt': 'SWITCH',
        'ups':   'UPS',        'sw':   'SOFTWARE',    'swk': 'SOFTWARE',
        'charm': 'CHARM_BASE', 'pwr':  'POWER',       'mc':  'MEDIA_CONV',
        'fopp':  'FOPP',       'fbr':  'FIBER',       'cab': 'CABINET',
        'mon':   'MONITOR',    'prt':  'PRINTER',     'con': 'CONSOLE',
        'fw':    'FIREWALL',   'scr':  'SCREEN',      'gen': 'SOFTWARE',
    }
    if cat_norm in cat_map:
        return cat_map[cat_norm], 'LOW'

    return 'UNKNOWN', 'UNKNOWN'


def classify_dataframe(df, rules=None):
    """
    Add 'diagram_class' and 'confidence' columns to the parsed BOM dataframe.

    Returns the dataframe with two new columns.
    """
    if rules is None:
        rules = load_rules()

    classes     = []
    confidences = []

    for _, row in df.iterrows():
        cls, conf = classify_row(
            description  = str(row.get('description', '')),
            part_number  = str(row.get('part_number', '')),
            category     = str(row.get('category', '')),
            area         = str(row.get('area', '')),
            rules        = rules
        )
        classes.append(cls)
        confidences.append(conf)

    df = df.copy()
    df['diagram_class'] = classes
    df['confidence']    = confidences
    return df


def save_user_correction(description: str, diagram_class: str,
                          rules_path=RULES_PATH):
    """
    When user manually corrects an UNKNOWN classification,
    save the description as a new keyword for that class.
    This is the self-learning mechanism.
    """
    rules = load_rules(rules_path)
    if diagram_class in rules:
        desc_norm = _normalize(description)
        existing  = [_normalize(k) for k in rules[diagram_class]['keywords']]
        if desc_norm not in existing:
            rules[diagram_class]['keywords'].append(desc_norm)
            with open(rules_path, 'w') as f:
                json.dump(rules, f, indent=2)
            return True
    return False


if __name__ == '__main__':
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from parser import parse_bom

    path = sys.argv[1] if len(sys.argv) > 1 else '../bom_full.csv'
    df   = parse_bom(path)
    df   = classify_dataframe(df)
    print(df[['description', 'area', 'diagram_class', 'confidence']].to_string())

    unknowns = df[df['diagram_class'] == 'UNKNOWN']
    if len(unknowns):
        print(f"\n⚠️  {len(unknowns)} UNKNOWN rows need manual classification:")
        print(unknowns[['description', 'category']].to_string())
    else:
        print("\n✅  All rows classified successfully.")
