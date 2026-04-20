# Planer Produkcji Okien Drewnianych (wersja wielouzytkowa)

Webowa aplikacja do planowania produkcji okien/drzwi oparta o:
- frontend: `index.html` + `app.js` + `styles.css`
- backend: `Flask` + `SQLite` (`planner.db`)

## Co jest w tej wersji

- Tryb wielouzytkowy: dane sa wspolne dla wszystkich klientow polaczonych z jednym serwerem.
- Kokpit: `Zamowienia`, `KPI`, `Gantt`, `Raporty`, `Wykonanie`, `Informacja zwrotna`, `Archiwum`, `Uzytkownicy`, `Ustawienia`.
- Planowanie terminu produkcji na podstawie:
  - czasow procesow (w minutach),
  - daty wejsciowej i dat dostepnosci materialow na poziomie pozycji,
  - checkboxow "Do zamowienia" dla materialow,
  - dni roboczych,
  - minut na zmiane oraz obsady stanowisk (zmiany i liczba osob per stanowisko).
- Statusy zamowien: Dokumentacja, Produkcja, Kosmetyka, Zakonczone.
- Status pozycji: aktualny dzial, aktualizowany recznie i masowo.
- Szczegoly zamowienia po kliknieciu wiersza + uwagi i zalacznik pozycji.
- Import zamowien i pozycji z pliku Excel (`.xlsx`) z poziomu sekcji `Zamowienia`.
- Wzorzec importu: `templates/import_zamowien_wzorzec.xlsx`.
- Ustawienia: zarzadzanie wieloma bazami danych (dodawanie pustych baz i przelaczanie aktywnej bazy; tylko admin).
- KPI z drill-down po kliknieciu kafelka (statusy/dzialy/materialy do zamowienia).
- Kokpit: ikona szybkiego dodawania zamowienia (okienko modalne) z kafelkiem przejscia do dodawania pozycji.
- Gantt dzienny zamowien + wejscie w zamowienie i podglad Gantta dla stanowisk.
- Raporty: wybor stanowisk i zakresu dat, lista zamowien/minut na dzien.
- Dni pracujace ustawiane indywidualnie (Pon-Ndz) w Ustawieniach.
- Technologie z indywidualnym procentowym rozpisaniem stanowisk per proces:
  - `IV68/78/92 Drewno`
  - `IV68/78/92 Drewno-Alu`
  - `Drzwi Drewno`, `Drzwi Drewno-Alu`
  - `HS`
  - `Otwierane na zewnatrz`
- Edycja procentow stanowisk z poziomu UI (Ustawienia -> Technologie i procenty stanowisk).
- Edycja ustawien stanowisk z poziomu UI (Ustawienia -> Obsada i zmiany na stanowiskach).

## Uruchomienie

1. Wejdz do katalogu projektu:
   - `cd C:\Users\piotr\Desktop\planer-test`
2. (Opcjonalnie) zainstaluj zaleznosci:
   - `python -m pip install -r requirements.txt`
3. (Opcjonalnie) skopiuj zmienne:
   - `copy .env.example .env`
4. Uruchom backend:
   - `python server.py`
5. Otworz aplikacje:
   - `http://127.0.0.1:8000`

## Logowanie i bezpieczenstwo

- API jest zabezpieczone sesja po stronie serwera (cookie).
- Po zalogowaniu sesja jest automatycznie przywracana po odswiezeniu strony.
- Do publikacji internetowej ustaw:
  - `APP_SECRET_KEY` (losowy, dlugi sekret)
  - `SESSION_COOKIE_SECURE=true` (dla HTTPS)

## Deploy na Render

W repo jest gotowy plik `render.yaml`.

1. Wrzuc projekt do GitHub.
2. W Render wybierz `New -> Blueprint` i wskaz repo.
3. Render utworzy usluge web + persistent disk wg `render.yaml`.
4. Po deployu otworz adres uslugi Render.

Uzyteczne zmienne:
- `PLANNER_DB_PATH=/opt/render/project/src/data/planner.db`
- `APP_SECRET_KEY=<losowy_sekret>`
- `SESSION_COOKIE_SECURE=true`

## Praca wielouzytkowa w sieci LAN

Jesli inni maja wejsc z innych komputerow w sieci lokalnej, otwieraja:
- `http://<IP_komputera_z_serwerem>:8000`

## Dane

- Baza: `planner.db` (SQLite) tworzona automatycznie.
- Warianty baz: katalog `variants/` (tworzone z panelu `Ustawienia`).
- Procenty technologii i wszystkie zamowienia sa zapisywane centralnie po stronie serwera.
