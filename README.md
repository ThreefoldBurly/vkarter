# vkarter
----

### WPROWADZENIE

**vkarter** to skrypt generujący *Karty Badań i Pomiarów Czynników Szkodliwych*. Obowiązek sporządzania takich kart nakłada na pracodawcę Kodeks Pracy, a ich wzór określony został w rozporządzeniu Ministra Zdrowia z dnia 2 lutego 2011 r. w sprawie badań i pomiarów czynników szkodliwych dla zdrowia w środowisku pracy [Dz.U. 2011 nr 33 poz. 166](http://isap.sejm.gov.pl/DetailsServlet?id=WDU20110330166 "Dz.U. 2011 nr 33 poz. 166").

**vkarter** generuje karty *z plików tekstowych DRK* specjalistycznego programu *HPZ99* służącego do wykonywania obliczeń badanych czynników szkodliwych i uciążliwych w środowisku pracy. *HPZ99* to program DOS-owy a mimo to ciągle szeroko stosowany w akredytowanych laboratoriach badań środowiska pracy, szczególnie na terenie Śląska.

Generowane karty to pliki MS Worda (.docx).

W katalogu `hpz` można znaleść 2 przykładowe zestawy plików DRK wraz z ich wydrukiem do PDF dla lepszego oglądu parsowanych danych.

### WYMAGANIA

Skrypt do poprawnego działania wymaga zainstalowania Pythona 2.7 oraz dwóch niestandardowych bibliotek:

- [python-docx](https://python-docx.readthedocs.io/en/latest/user/install.html "python-docx")
- [FuzzyWuzzy](https://github.com/seatgeek/fuzzywuzzy "FuzzyWuzzy")

Powinno się dać go uruchomić na każdej platformie (Linux/Windows/MacOS), ale nie daję głowy - sam przetestowałem tylko na Ubuntu.

### UŻYCIE

Przegrać pliki .py i .vka do katalogu, zawierającego pliki DRK, z których chcemy wygenerować karty. W terminalu, będąc w tym katalogu, uruchomić skrypt komendą:

`python vkarter_output.py parametry`

Skrypt rozpoznaje 3 scenariusze parametrów:

1) *parametry* to nazwy plików do sparsowania oddzielone spacją
2) *parametry* to `-a` albo `--all` - zostaną wczytane wszystkie pliki DRK znajdujące się w katalogu
3) *parametry* to `-a 00-000` lub `--all 00-000`, gdzie `00-000` to nazwa pliku DRK w formacie *rok-nr_sprawozdania* np. `17-205.DRK` - wszystkie pliki w takim formacie (np. `17-205.DRK`, `17-205-2.DRK`, `17-205-3.DRK` itd.) zostaną wczytane

##### PLIKI TEKSTOWE

`vkarter_stale.vka` - tu można łatwo zmienić wykonawcę lub metody pomiarów

`vkarter_skroty.vka` - tu można dodawać skróty używane przy tworzeniu nazw generowanych plików

`vkarter_odmiany.vka` - tu należy dodać odmianę związku chemicznego, którego **vkarter** nie rozpoznał

Jeśli ktoś woli DOC zamiast DOCX, to polecam pakiet `unoconv` (trzeba mieć zainstalowanego LibreOffice'a) i komendę:
`unoconv -f doc *.docx; rm ./*.docx`

### ZASTRZEŻENIA

**vkarter** to WIP. W założeniach miał również generować *Rejestry Czynników Szkodliwych dla Zdrowia Występujących na Stanowisku Pracy** (parsuje wszystkie dane, które są do tego potrzebne). Z czynników szkodliwych oprócz hałasu, pyłów i substancji chemicznych miał też brać pod uwagę drgania - tu praca utknęła w pół drogi. Do tego program ma kilka bugów, o których wiem i których nigdy nie wyprostowałem (i pewno więcej, o których nie wiem). 

Poza tym główna biblioteka, którą się posłużyłem, `python-docx` sama jest WIP i narzuca pewne ograniczenia, których się nie da obejść.



