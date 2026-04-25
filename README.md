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

## Co nowego w v13 (odświeżenie UI)
- nowy system typograficzny: **Onest** (interfejs) + **JetBrains Mono** (numery, kody, ceny)
- wskaźnik kroków w nagłówku — widać w jakim etapie jesteś i które kroki są zakończone
- spójna paleta kolorów: ciepłe, papierowe tło zamiast zimnego niebieskawego
- wszystkie karty i tabele wyrównane symetrycznie, jednakowe paddingi i marginesy
- czytelniejsze pigułki z producentem, kodem własnym i EAN-em
- pasek akcji na dole zmienił się w premium kontrolkę z czarnym tłem i wyróżnionym przyciskiem eksportu
- liczby i kody wyświetlane czcionką tabular-nums (kolumny się idealnie wyrównują)
- na desktopie lista refundacji rozkłada się na 2 kolumny — mniej scrolla
- delikatna ziarnistość tła i animacje wejścia sekcji dla głębi

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
3. po odświeżeniu aplikacji kliknij `Odśwież`

## GitHub Pages
1. wrzuć cały folder projektu do repozytorium
2. w ustawieniach repo włącz GitHub Pages
3. jako źródło wybierz branch z projektem
4. po publikacji aplikacja będzie działała jako strona i jako PWA

## Instalacja
Na telefonie lub komputerze w Chrome albo Edge pojawi się opcja instalacji.

## Aktualizacja po deployu
Po wgraniu nowej wersji v13 service worker automatycznie podmieni cache (zmiana wersji z `refundacje-pwa-v12` na `refundacje-pwa-v13`). Jeśli nie widać nowego wyglądu — odśwież stronę dwa razy lub użyj "Hard reload" (Ctrl/Cmd+Shift+R).

## Uwaga
Obsługa plików `xls/xlsx` działa w przeglądarce. Przy pierwszym otwarciu aplikacja powinna mieć dostęp do internetu, żeby załadować bibliotekę do odczytu i zapisu plików Excel oraz fonty.
