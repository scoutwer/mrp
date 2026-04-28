"""
mrp_logic.py — Czysta logika algorytmu MRP (Material Requirements Planning).

Moduł niezależny od GUI. Zawiera:
- Obliczanie tabeli MRP (zapotrzebowanie netto, planowane zamówienia, itp.)
- Eksplozja BOM (wyliczanie zapotrzebowania dziecka na podstawie zamówień rodzica)
"""

import math
from dataclasses import dataclass, field


@dataclass
class MRPResult:
    """Wynik obliczeń algorytmu MRP dla jednego indeksu materiałowego."""
    proj_avail: list = field(default_factory=list)
    net_req: list = field(default_factory=list)
    planned_order_rec: list = field(default_factory=list)
    planned_order_rel: list = field(default_factory=list)
    bledy: list = field(default_factory=list)


def oblicz_mrp(
    na_stanie: int,
    czas_realizacji: int,
    wielkosc_partii: int,
    gross_req: list,
    sched_rec: list,
    n: int,
) -> MRPResult:
    """
    Oblicza pełną tabelę MRP.

    Args:
        na_stanie: Zapas początkowy (stan magazynowy na początku horyzontu).
        czas_realizacji: Lead Time — ile okresów wcześniej trzeba złożyć zamówienie.
        wielkosc_partii: Wielkość partii (0 = lot-for-lot / partia na partię).
        gross_req: Lista całkowitego zapotrzebowania brutto (per okres).
        sched_rec: Lista planowanych przyjęć (scheduled receipts, per okres).
        n: Liczba okresów (tygodni) w horyzoncie planowania.

    Returns:
        MRPResult z wypełnionymi listami wyników i ewentualnymi błędami.
    """
    proj_avail = [0] * n
    net_req = [0] * n
    planned_order_rec = [0] * n
    planned_order_rel = [0] * n
    bledy = []

    for i in range(n):
        zapas_poprzedni = proj_avail[i - 1] if i > 0 else na_stanie
        stan_przed = zapas_poprzedni + sched_rec[i] - gross_req[i]

        if stan_przed < 0:
            brakuje = abs(stan_przed)
            net_req[i] = brakuje

            # Wyznacz wielkość zamówienia (lot sizing)
            if wielkosc_partii > 0:
                mnozenie = math.ceil(brakuje / wielkosc_partii)
                zamowienie = mnozenie * wielkosc_partii
            else:
                zamowienie = brakuje

            # Przesunięcie zamówienia w czasie o Lead Time
            rel_idx = i - czas_realizacji

            if rel_idx >= 0:
                planned_order_rel[rel_idx] = zamowienie
                planned_order_rec[i] = zamowienie
                proj_avail[i] = stan_przed + zamowienie
            else:
                # Zamówienie wypada przed horyzontem — brak czasu na realizację
                bledy.append(
                    f"⚠ BŁĄD LOGISTYCZNY (Tydzień {i + 1}): "
                    f"Brak czasu na zamówienie! Zapas spada do {stan_przed} szt. "
                    f"(wymagane zamówienie w tygodniu {rel_idx + 1}, przed horyzontem)"
                )
                planned_order_rec[i] = 0
                proj_avail[i] = stan_przed
        else:
            net_req[i] = 0
            planned_order_rec[i] = 0
            proj_avail[i] = stan_przed

    return MRPResult(
        proj_avail=proj_avail,
        net_req=net_req,
        planned_order_rec=planned_order_rec,
        planned_order_rel=planned_order_rel,
        bledy=bledy,
    )


def oblicz_zapotrzebowanie_bom(
    zamowienia_rodzica: list,
    mnoznik: int,
    n: int,
) -> list:
    """
    Eksplozja BOM — wylicza Całkowite Zapotrzebowanie dziecka
    na podstawie Planowanych Zamówień (rel) rodzica pomnożonych przez mnożnik.

    Args:
        zamowienia_rodzica: Lista 'planned_order_rel' rodzica.
        mnoznik: Ile sztuk dziecka potrzeba na 1 sztukę rodzica.
        n: Liczba okresów.

    Returns:
        Lista zapotrzebowania brutto dziecka (długość n).
    """
    wynik = []
    for i in range(n):
        if i < len(zamowienia_rodzica):
            wynik.append(zamowienia_rodzica[i] * mnoznik)
        else:
            wynik.append(0)
    return wynik


def bezpieczny_int(wartosc: str) -> int:
    """Konwertuje string na int, zwraca 0 dla pustych/nieprawidłowych wartości."""
    wartosc = wartosc.strip()
    if not wartosc:
        return 0
    try:
        return int(wartosc)
    except ValueError:
        try:
            return int(float(wartosc))
        except ValueError:
            return 0
