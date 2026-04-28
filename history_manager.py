"""
history_manager.py — Zarządzanie historią planów MRP (plik JSON).

Moduł odpowiada za:
- Odczyt historii z pliku JSON
- Zapis/aktualizację profili
- Usuwanie profili
- Listowanie dostępnych kluczy
"""

import json
import os


class HistoryManager:
    """Menedżer historii planów MRP przechowywanych w pliku JSON."""

    def __init__(self, plik_historii: str = "mrp_history.json"):
        self.plik_historii = plik_historii
        self._dane: dict = {}

    # ------------------------------------------------------------------
    # Odczyt
    # ------------------------------------------------------------------

    def wczytaj(self) -> dict:
        """
        Wczytuje historię z pliku JSON.
        Jeśli plik nie istnieje lub jest uszkodzony, zwraca pusty słownik.
        """
        if os.path.exists(self.plik_historii):
            try:
                with open(self.plik_historii, "r", encoding="utf-8") as f:
                    self._dane = json.load(f)
            except (json.JSONDecodeError, IOError):
                self._dane = {}
        else:
            self._dane = {}
        return self._dane

    @property
    def dane(self) -> dict:
        """Aktualny stan danych w pamięci."""
        return self._dane

    # ------------------------------------------------------------------
    # Zapis
    # ------------------------------------------------------------------

    def zapisz(self) -> None:
        """Zapisuje aktualny stan danych do pliku JSON."""
        with open(self.plik_historii, "w", encoding="utf-8") as f:
            json.dump(self._dane, f, indent=4, ensure_ascii=False)

    def dodaj(self, nazwa: str, dane_mrp: dict) -> None:
        """
        Dodaje lub nadpisuje profil w historii i zapisuje do pliku.

        Args:
            nazwa: Klucz profilu (nazwa detalu).
            dane_mrp: Słownik z pełnymi danymi MRP.
        """
        self._dane[nazwa] = dane_mrp
        self.zapisz()

    # ------------------------------------------------------------------
    # Usuwanie
    # ------------------------------------------------------------------

    def usun(self, nazwa: str) -> bool:
        """
        Usuwa profil z historii.

        Args:
            nazwa: Klucz profilu do usunięcia.

        Returns:
            True jeśli usunięto, False jeśli klucz nie istniał.
        """
        if nazwa in self._dane:
            del self._dane[nazwa]
            self.zapisz()
            return True
        return False

    # ------------------------------------------------------------------
    # Pomocnicze
    # ------------------------------------------------------------------

    def lista_kluczy(self) -> list:
        """Zwraca listę nazw zapisanych profili."""
        return list(self._dane.keys())

    def pobierz(self, nazwa: str) -> dict | None:
        """Zwraca dane profilu lub None jeśli nie istnieje."""
        return self._dane.get(nazwa)

    def czy_istnieje(self, nazwa: str) -> bool:
        """Sprawdza czy profil o podanej nazwie istnieje."""
        return nazwa in self._dane
