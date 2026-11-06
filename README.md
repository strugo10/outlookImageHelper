# Outlook Image Helper (Outlook - Pomocnik obrazów i OCR)

Dodatek dla czytnika ekranu **NVDA**, który zapewnia pełną dostępność, rozpoznawanie tekstu (OCR), interaktywną rozmowę i opisywanie przez sztuczną inteligencję (**Google Gemini**), bibliotekę promptów, obsługę linków oraz inteligentny i seryjny zapis grafik wklejonych w wiadomości **Microsoft Outlook Classic**.

---

## ✨ Kluczowe możliwości dodatku

1. **Menu czynności (kontekstowe) na liście grafik:**
   - Klawisz `Aplikacje`, `Shift+F10` lub przycisk `Alt+C` na dowolnej grafice otwiera pełne menu dostępnych operacji.
   - Gdy zaznaczysz kilka grafik spacją (`Spacja` / `Shift+strzałki`), menu automatycznie przełącza się na akcje zbiorcze (zbiorczy OCR, zbiorczy opis AI, zapisanie zaznaczonych do folderu).

2. **Własne prompty i biblioteka szablonów AI:**
   - Możliwość wyboru gotowego szablonu instrukcji (np. *Tabela kadrowa / urlopy*, *Okno programu / błędy*, *Wykresy*, *Dokładne przepisanie tekstu*).
   - Opcja tworzenia, zapisywania i usuwania własnych szablonów oraz jednorazowej modyfikacji prompta przed wysłaniem.

3. **Interaktywna rozmowa z AI (Q&A o grafikę):**
   - Po wstępnym opisie grafiki otwiera się okno rozmowy wieloturowej.
   - Możliwość dopytywania AI o szczegóły tabeli, kwoty, nazwiska, daty czy komunikaty o błędach (`Ctrl+Enter` wysyła pytanie).
   - AI pamięta całą historię rozmowy oraz obraz.

4. **Uniwersalny format kopiowania i eksportu z grafiką:**
   - Przyciski *Kopiuj ostatnią odpowiedź* oraz *Kopiuj całą rozmowę*.
   - Checkbox *Dołącz źródłową grafikę*: umieszcza w schowku bogaty format HTML ze zintegrowanym obrazem Base64. Wklejony do Outlooka lub Worda zachowuje obraz oraz sformatowany tekst! W Notatniku wkleja czysty tekst.
   - Eksport raportu do samodzielnego pliku `.html` lub dokumentu `.md` (Markdown).

5. **Wykrywanie i otwieranie hiperłączy z grafik (YouTube, LinkedIn itp.):**
   - Dodatek wykrywa odnośniki wklejone w grafiki i prezentuje je na liście (`[Link: youtube.com]`).
   - Bezpośrednie otwarcie w przeglądarce przez menu czynności, przycisk lub skrót `Ctrl+Enter`.

---

## ⌨️ Główne skróty klawiszowe

- `Ctrl+NVDA+A` (lub `NVDA+Shift+G`, `Alt+Shift+G`) – Otwiera okno dialogowe „Lista grafik w wiadomości”.
- `NVDA+Windows+G` lub `Alt+Shift+O` – Szybki OCR bieżącej grafiki lub zrzutu w mailu.
- `NVDA+Windows+Shift+G` lub `Alt+Shift+A` – Szybki opis bieżącej grafiki przez AI.
- W oknie listy grafik:
  - `Aplikacje` / `Shift+F10` / `Alt+C` – Menu czynności.
  - `Ctrl+Enter` – Otwarcie linku przypisanego do grafiki w przeglądarce.
  - `Spacja` / `Shift+strzałki` – Zaznaczanie wielu grafik do operacji seryjnych.

---

## 🛠️ Budowanie paczki instalacyjnej

Aby zbudować paczkę `.nvda-addon`:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
```
