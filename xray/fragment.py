"""Aperçu XML d'une cellule : la balise CII que produit sa valeur (« le BT se transforme en balise »).

Aperçu pédagogique construit à partir du chemin XML et de la valeur saisie, avec la même mise en forme
que la facture générée (generation/formats.py) : il montre l'élément, pas le document complet.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape, quoteattr

from generation.formats import formater

ROOT = "/rsm:CrossIndustryInvoice/"


def _segments(xpath: str) -> list[str]:
    """Découpe sur « / » hors des crochets."""
    parts, depth, current = [], 0, ""
    for ch in xpath[len(ROOT):] if xpath.startswith(ROOT) else xpath.lstrip("/"):
        depth += (ch == "[") - (ch == "]")
        if ch == "/" and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += ch
    return parts + [current] if current else parts


def _parse(segment: str):
    """« ram:ID[@schemeID='VA'] » → (balise, attributs, enfant imposé)."""
    m = re.match(r"^([\w:]+)(?:\[(.*)\])?$", segment)
    tag, predicate = m.group(1), m.group(2)
    attrs, child = {}, None
    if predicate:
        a = re.match(r"^@(\w+)='([^']*)'$", predicate)
        c = re.match(r"^([\w:]+)='([^']*)'$", predicate)
        if a:
            attrs[a.group(1)] = a.group(2)
        elif c:
            child = (c.group(1), c.group(2))
    return tag, attrs, child


def fragment(xpath: str, value, type_valeur: str = "texte") -> str:
    """Fragment XML indenté, du premier élément sous la racine jusqu'à la valeur."""
    segments = _segments(xpath)
    attribut = segments.pop()[1:] if segments and segments[-1].startswith("@") else None
    texte = formater(type_valeur, value) or "…"
    parsed = [_parse(s) for s in segments]
    lines, closing = [], []
    for depth, (tag, attrs, child) in enumerate(parsed):
        pad = "  " * depth
        last = depth == len(parsed) - 1
        if last and attribut:
            attrs = {**attrs, attribut: texte}
        elif last and type_valeur == "date":
            attrs = {**attrs, "format": "102"}
        attr_text = "".join(f" {k}={quoteattr(v)}" for k, v in attrs.items())
        if last:
            lines.append(f"{pad}<{tag}{attr_text}>{'…' if attribut else escape(texte)}</{tag}>")
        else:
            lines.append(f"{pad}<{tag}{attr_text}>")
            closing.append((pad, tag, child))
    for pad, tag, child in reversed(closing):
        if child:  # ex. AdditionalReferencedDocument[ram:TypeCode='916'] : le code accompagne la référence
            lines.append(f"{pad}  <{child[0]}>{escape(child[1])}</{child[0]}>")
        lines.append(f"{pad}</{tag}>")
    return "\n".join(lines)
