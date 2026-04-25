# Generator refundacji

Gotowa aplikacja PWA do wrzucenia na GitHub Pages.

## Co robi
- wczytuje dane tylko z folderu `data`
- obsługuje pliki `klienci.xlsx`, `klienci.xls` albo `klienci.json`
- obsługuje pliki `towar.xlsx`, `towar.xls` albo `towar.json`
- kolejność szukania jest taka: `.xlsx` → `.xls` → `.json`
- pozwala wybrać wielu kontrahentów jednocześnie
- wyszukuje kontrahentów po połączonym polu: `Numer | Logo | Nazwa`
- filtruje towary po producencie `NazwaZnacznika`
- wyszukuje indeksy po `KodWlasny` i `Nazwa`
- dodaje indeksy do obszaru roboczego pojedynczo albo zbiorczo z widocznej listy
- pozwala ustawić cenę AC i refundację osobno dla każdego indeksu
- pozwala zaznaczyć wiele indeksów i wykonać zbiorczą edycję ceny AC oraz refundacji
- dodaje wspólny zakres dat refundacji `od` i `do`
- eksport zawiera kolumnę `cena AC` przed kolumną `refundacja`
- eksportuje gotowy plik Excel `.xlsx`
- ma układ uproszczony pod telefon: jedna aktywna sekcja, mniej przewijania, dolny pasek eksportu

## Pliki w folderze data
Możesz trzymać tam jedną z poniższych wersji:

### Kontrahenci
- `data/klienci.xlsx`
- `data/klienci.xls`
- `data/klienci.json`

Wymagane kolumny:
- `Numer`
- `Logo`
- `Nazwa`

### Towary
- `data/towar.xlsx`
- `data/towar.xls`
- `data/towar.json`

Wymagane kolumny:
- `NazwaZnacznika`
- `KodWlasny`
- `Nazwa`

## Jak aktualizować dane
1. podmień plik w folderze `data`
2. zacommituj zmiany do repozytorium
3. po odświeżeniu aplikacji kliknij `Odśwież dane`

## GitHub Pages
1. wrzuć cały folder projektu do repozytorium
2. w ustawieniach repo włącz GitHub Pages
3. jako źródło wybierz branch z projektem
4. po publikacji aplikacja będzie działała jako strona i jako PWA

## Instalacja
Na telefonie lub komputerze w Chrome albo Edge pojawi się opcja instalacji.

## Uwaga
Obsługa plików `xls/xlsx` działa w przeglądarce. Przy pierwszym otwarciu aplikacja powinna mieć dostęp do internetu, żeby załadować bibliotekę do odczytu i zapisu plików Excel.
