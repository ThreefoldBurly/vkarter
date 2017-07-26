# -*- coding: utf-8 -*-
from sys import argv, exit
from os import getcwd, listdir
from os.path import exists
import codecs

from vkarter_shared import *
from vkarter_czynniki import *


class ParserDRK(object):
    """
    Parser plików DRK. Pobiera i stronicuje pliki DRK, po czym łączy strony w Pomiary (porcje gotowe do dalszego parsowania) i grupuje wg 3 głównych typów, w jakich występują pliki DRK
    """
    def __init__(self):
        super(ParserDRK, self).__init__()
        # TODO: sprawdzić sensowność istnienia pola poniżej
        self.nazwy_plikow = []  # lista nazw wczytanych plików
        self.pliki_DRK = self.pobierzDRK()  # lista zawartości każdego z wczytanych plików (lista list linii)
        self.strony_DRK = self.stronicujDRK()  # lista stron wszystkich wczytanych plików DRK
        self.pomiary = self.grupujPomiary()  # obiekty klasy Pomiar uporządkowane w GlTypyPomiarow

    @staticmethod
    def poprawNazwy():
        """Dodaje do nazw plików odpowiednie rozszerzenie (jeśli trzeba)"""
        for i in xrange(1, len(argv)):
            nazwa = argv[i]
            rozsz = nazwa[-4:]

            if rozsz.lower() != ".drk":
                nazwa = nazwa + ".drk"
                argv[i] = nazwa

    @staticmethod
    def sprawdzWzorzecNazwy(wzorzec):
        """
        Sprawdza czy nazwa pliku do wczytania jest zgodna z formatem XX-XXX, gdzie X do liczba
        """
        if len(wzorzec) == 6:
            if wzorzec[:2].isdigit() and wzorzec[2] == "-" and wzorzec[-3:].isdigit():
                return True
        else:
            return False

    def pobierzDRK(self):
        """Czyta pliki DRK, zwraca listę ich zawartości (każda zawartość to lista linii)"""
        pliki_DRK = []

        if len(argv) > 1:
            pierwszy_arg = argv[1]

            # 1) najpierw obsługa parametru w formacie "-a" lub "--all" rozumianego jako: wczytaj wszystkie pliki DRK znajdujące się w aktualnym katalogu
            if len(argv) == 2 and (pierwszy_arg == "-a" or pierwszy_arg == "--all"):
                lista_plikow = listdir(getcwd())
                if len(lista_plikow) > 0:
                    for plik in lista_plikow:
                        rozsz = plik[-4:]
                        if rozsz.lower() == ".drk":
                            # WAŻNE: here comes Unicode! Ta linia (i jej powtórzenia) zamienia bajty kodowane windowsowym 'cp1250' z pliku na obiekty (stringi) klasy unicode
                            zawartosc = codecs.open(plik, encoding='cp1250').readlines()
                            pliki_DRK.append(zawartosc[5:])  # pierwsze pięć linii nie ma znaczenia, więc jest obcinane
                            self.nazwy_plikow.append(plik)
                            print "Załadowano plik: ", plik  # test

                    if len(pliki_DRK) == 0:
                        print "Nie znaleziono pliku(ów) DRK do załadowania"
                        exit(0)
                else:
                    print "Katalog nie zawiera żadnych plików"

            # 2) teraz obsługa parametru w formacie "-a 00-000" lub "--all 00-000" rozumianego jako: wczytaj wszystkie DRK dla podanego numeru zlecenia
            elif len(argv) == 3 and (pierwszy_arg == "-a" or pierwszy_arg == "--all") and self.sprawdzWzorzecNazwy(argv[2]):
                lista_plikow = listdir(getcwd())
                if len(lista_plikow) > 0:
                    for plik in lista_plikow:
                        rozsz = plik[-4:]
                        if rozsz.lower() == ".drk" and plik[:6] == argv[2]:
                            zawartosc = codecs.open(plik, encoding='cp1250').readlines()
                            pliki_DRK.append(zawartosc[5:])
                            self.nazwy_plikow.append(plik)
                            print "Załadowano plik: ", plik  # test

                    if len(pliki_DRK) == 0:
                        print "Nie znaleziono pliku(ów) DRK o podanej nazwie"
                        exit(0)
                else:
                    print "Katalog nie zawiera żadnych plików"
                    exit(0)

            # 3) na końcu obsługa sytuacji, w której każdy parametr to nazwa pliku do załadowania
            else:
                self.poprawNazwy()

                for nazwa in argv[1:]:
                    rozsz = nazwa[-4:]
                    nazwa_alt = nazwa[:-4] + rozsz.upper()

                    if exists(nazwa):
                        zawartosc = codecs.open(nazwa, encoding='cp1250').readlines()
                        pliki_DRK.append(zawartosc[5:])
                        self.nazwy_plikow.append(nazwa)
                        print "Załadowano plik: ", nazwa  # test
                    elif exists(nazwa_alt):
                        zawartosc = codecs.open(nazwa_alt, encoding='cp1250').readlines()
                        pliki_DRK.append(zawartosc[5:])
                        self.nazwy_plikow.append(nazwa_alt)
                        print "Załadowano plik: ", nazwa_alt  # test
                    else:
                        print "Nie mogę otworzyć pliku o nazwie: ", nazwa
                        exit(0)
        else:
            print "Nie podano pliku(ów) DRK do opracowania"
            exit(0)

        return pliki_DRK

    def stronicujDRK(self):
        """
        Dzieli przeczytane zawartości plików DRK na strony i zwraca jako listę stron (każda strona to lista linii)
        """
        strony = []

        for plik in self.pliki_DRK:
            jedna_strona = []

            for linia in plik:
                if u"" not in linia:

                    # obsługa ostatniej strony
                    if u"" in linia and len(linia) == 1:
                        strony.append(jedna_strona)
                        jedna_strona = []
                        continue

                    jedna_strona.append(linia)

                    # wycięcie pustej linii która znajduje się na początku każdej kolejnej strony poza pierwszą
                    if len(jedna_strona) == 1 and linia == '\r\n':
                        jedna_strona.pop()

                else:
                    strony.append(jedna_strona)
                    jedna_strona = []

        return strony

    def grupujPomiary(self):
        """
        Grupuje wszystkie strony wczytanych plików DRK w Pomiary z podziałem na 3 główne typy. Ze zgrupowanych Pomiarów tworzy obiekty klasy Pomiar i zwraca z podziałem na typy (jako krotkę GlTypyPomiarow)
        """

        halas_DRK = []
        drgania_DRK = []
        pyly_chemia_DRK = []

        # najpierw spłaszczenie głównej listy (moje pierwsze list comprehension!)
        # (już niepotrzebne)
        # wszystkie_strony = [strona for strony in zestr_pliki_DRK for strona in strony]

        # oznaczenie stron znacznikami typu Pomiaru
        oznaczenia_stron = []
        znaczniki_pyl_chem = ZNACZNIKI_POM[2:]  # znaczniki dla pyłów, chemii, pyłów i chemii

        for strona in self.strony_DRK:
            for linia in strona:
                if ZNACZNIKI_POM.halas in linia:
                    oznaczenia_stron.append(TypPomiaru.HALAS)
                    break
                elif ZNACZNIKI_POM.drgania in linia:
                    oznaczenia_stron.append(TypPomiaru.DRGANIA)
                    break
                elif any(znacznik in linia for znacznik in znaczniki_pyl_chem):  # skrócony zapis 3 alternatyw
                    oznaczenia_stron.append(TypPomiaru.PYLY_CHEMIA)
                    break
            else:  # nie znaleziono znacznika Pomiarów
                oznaczenia_stron.append(TypPomiaru.ZADEN)

        # grupowanie
        pomiar_DRK = []  # lista linii
        ost_typ_pomiaru = TypPomiaru.ZADEN  # wartość startowa

        for strona, oznaczenie in zip(self.strony_DRK, oznaczenia_stron):

            if oznaczenie != TypPomiaru.ZADEN:
                # ta strona zawiera nowy Pomiar, więc najpierw kończymy zapis starego...
                if len(pomiar_DRK) > 0:  # w pierwszej iteracji nie ma jeszcze starego Pomiaru do zapisu
                    if ost_typ_pomiaru == TypPomiaru.HALAS:
                        halas_DRK.append(pomiar_DRK)
                    elif ost_typ_pomiaru == TypPomiaru.DRGANIA:
                        drgania_DRK.append(pomiar_DRK)
                    elif ost_typ_pomiaru == TypPomiaru.PYLY_CHEMIA:
                        pyly_chemia_DRK.append(pomiar_DRK)

                pomiar_DRK = []

                # ...a potem rozpoczynamy zapis nowego
                for linia in strona:
                    pomiar_DRK.append(linia)

                ost_typ_pomiaru = oznaczenie

            else:
                # ta strona jest tylko kontynuacją Pomiaru, więc dodaj
                for linia in strona:
                    pomiar_DRK.append(linia)

        # obsługa ostatniego Pomiaru
        if ost_typ_pomiaru == TypPomiaru.HALAS:
            halas_DRK.append(pomiar_DRK)
        elif ost_typ_pomiaru == TypPomiaru.DRGANIA:
            drgania_DRK.append(pomiar_DRK)
        elif ost_typ_pomiaru == TypPomiaru.PYLY_CHEMIA:
            pyly_chemia_DRK.append(pomiar_DRK)

        # kreacja obiektów klasy Pomiar ze zgrupowanych pomiarów DRK
        halas = []
        for i in xrange(len(halas_DRK)):
            pomiar_DRK = halas_DRK[i]
            halas.append(PomiarHalas(i + 1, "H", pomiar_DRK))  # numer Pomiaru nie jest indeksem listy

        drgania = []
        for i in xrange(len(drgania_DRK)):
            pomiar_DRK = drgania_DRK[i]
            drgania.append(PomiarDrgania(i + 1, "D", pomiar_DRK))

        pyly_chemia = []
        for i in xrange(len(pyly_chemia_DRK)):
            pomiar_DRK = pyly_chemia_DRK[i]
            pyly_chemia.append(PomiarPylyChemia(i + 1, "PCH", pomiar_DRK))

        return GlTypyPomiarow(halas, drgania, pyly_chemia)

    # getter dla klas testujących
    def podajPomiar(self, typ, ktory_pomiar=1):
        """Podaje wybrany Pomiar z listy Pomiarów danego typu . Domyślnie podaje pierwszy Pomiar"""
        lista_pomiarow = self.pomiary._asdict()[typ]

        assert len(lista_pomiarow) > 0, ("Podano niewlasciwy typ Pomiarów do wyswietlenia: " + typ)
        assert ktory_pomiar in range(1, len(lista_pomiarow) + 1), "Podany numer Pomiaru poza zakresem."

        return lista_pomiarow[ktory_pomiar - 1]


class Pomiar(object):
    """
    Abstrakcyjna reprezentacja Pomiaru (porcji pliku DRK odpowiadającej 1 pomiarowi (jednego lub wielu czynników))
    """
    def __init__(self, numer, typ, tresc):
        super(Pomiar, self).__init__()
        self.numer = numer  # przydatne do debuggowania
        self.typ = typ  # przydatne do debuggowania
        self.tresc = self.poprawTresc(tresc)  # lista linii

    @staticmethod
    def poprawTresc(tresc):
        """
        Usuwa z treści puste linie, z każdej niepustej linii usuwa niepotrzebne znaki (1 na początku i 2 na końcu)
        """
        nowa_tresc = []
        for i in xrange(len(tresc)):
            linia = tresc[i]
            if i == 0:  # pierwsza linia jest inna niż reszta (nie ma znaku do wycięcia na początku)
                nowa_tresc.append(linia[0:-2])
            else:
                if len(linia) > 2:
                    nowa_tresc.append(linia[1:-2])

        return nowa_tresc


class NaglowekPomiaru(object):
    """Wrapper dla (pierwszego) naglowka Pomiaru. Zapewnia parserom wygodny dostęp do sekcji"""
    def __init__(self, tresc):
        super(NaglowekPomiaru, self).__init__()
        self.sekcja_miejsca_i_nazwy, self.sekcja_daty = self.wydzielSekcje(tresc)

    @staticmethod
    def wydzielSekcje(tresc):
        """Z nagłówka Pomiaru wydziela i zwraca sekcję miejsca i nazwy oraz sekcję daty"""
        sekcja_miejsca_i_nazwy = []  # lista linii
        sekcja_daty = []  # lista linii
        druga_sekcja = False

        for linia in tresc[1:]:  # pomijamy pierwszą linie
            if not druga_sekcja:
                if not any(znacznik in linia for znacznik in ZNACZNIKI_POM):
                    sekcja_miejsca_i_nazwy.append(linia)
                else:
                    druga_sekcja = True
                    continue
            else:
                sekcja_daty.append(linia)

        if len(sekcja_miejsca_i_nazwy) == 0:
            sekcja_miejsca_i_nazwy = None

        return (sekcja_miejsca_i_nazwy, sekcja_daty)


class TabelaPomiaru(object):
    """Wrapper dla tabeli Pomiaru. Zapewnia parserom wygodny dostęp do wierszy"""
    def __init__(self, tresc):
        super(TabelaPomiaru, self).__init__()
        self.wiersze = self.wydzielWiersze(tresc)

    @staticmethod
    def wydzielWiersze(tresc):
        """Z tabeli Pomiaru wydziela i zwraca wiersze"""
        wiersze = []

        wiersz = []
        nr_wiersza = 1
        for i in xrange(len(tresc)):
            linia = tresc[i]

            if linia[0] == u"Ě":
                # warunek dla pominięcia pośrednich poziomych ramek w tabeli NDSch
                # if len(linia) > 19: - odrzucony po testach
                # if linia[-2:] != u".Ď": - odrzucony po testach
                if u"ÍÉHÎÉ0ÎÉ)" not in linia:
                    wiersz.append(linia)
            else:
                if i != 0:
                    wiersze.append(WierszTabeliPomiaru(nr_wiersza, wiersz))
                    nr_wiersza += 1
                    wiersz = []
                    continue

        return wiersze


class WierszTabeliPomiaru(object):
    """Wrapper dla wiersza tabeli Pomiaru. Zapewnia parserom wygodny dostęp do kolumn"""
    def __init__(self, numer, tresc):
        super(WierszTabeliPomiaru, self).__init__()
        self.numer = numer
        self.tresc = tresc  # lista linii
        # UWAGA! Maksymalna ilość kolumn dla typów Pomiarów (wynikająca z budowy tabeli):
        # HAŁAS: 10
        # DRGANIA: 3
        # PYŁY SiO2: 7
        # POZOSTAŁE: 6
        self.kolumny = self.wydzielKolumny()  # lista linii

    def wydzielKolumny(self):
        rzedy = []

        for linia in self.tresc:
            porcje = linia.split(u'Ě')[1:-1]
            porcje_popr = []

            for porcja in porcje:
                # wycięcie pierwszego znaku (spacja) i ostatniego znaku z każdej porcji (śmieć)
                porcja = porcja[1:-1]
                # wycięcie ostatniego znaku, jeśli takowy istnieje i jest spacją
                if len(porcja) > 0 and porcja[-1].isspace():
                    porcja = porcja[:-1]
                porcje_popr.append(porcja)
            rzedy.append(porcje_popr)

        return zip(*rzedy)

    # TYMCZASOWO - METODY DO TESTÓW
    def wyswietlTresc(self):
        komunikat = "Wiersz #%d tabeli Pomiaru:" % self.numer
        wyswietlKomunikat(komunikat)
        for linia in self.tresc:
            print linia

    def wyswietlKolumne(self, nr_kolumny):
        komunikat = "Kolumna #%d wiersza #%d tabeli Pomiaru:" % (nr_kolumny, self.numer)
        wyswietlKomunikat(komunikat)
        for rzad in self.kolumny[nr_kolumny - 1]:
            print repr(rzad)


class PomiarHalas(Pomiar):
    """Reprezentacja Pomiaru I typu uwzględniająca podział na sekcje właściwe temu typowi."""
    def __init__(self, numer, typ, tresc):
        super(PomiarHalas, self).__init__(numer, typ, tresc)
        # sekcje
        self.naglowek, self.tabela, self.stopka = self.wydzielSekcje()

    def wydzielSekcje(self):
        """Dzieli treść Pomiaru na sekcje: nagłówek, tabelę, oraz stopkę i zwraca je w krotce"""
        naglowek = []
        tabela = []
        stopka = []
        # flagi
        jest_naglowek = True
        jest_tabela = False
        dodawaj = False
        poczatek_tab = True  # flaga konieczna dla prawidłowej obsługi linii rozdzielających wiersze tabeli

        for linia in self.tresc:
            # NAGŁÓWEK
            if jest_naglowek:
                if linia[0] != u"Č":
                    naglowek.append(linia)
                else:
                    jest_naglowek = False
                    jest_tabela = True
                    continue

            # TABELA
            elif jest_tabela:
                # obsługa początku tabeli i końca przerwy
                if not dodawaj:
                    if linia[0] == u"Í":
                        dodawaj = True
                        if poczatek_tab:
                            tabela.append(linia)
                            poczatek_tab = False
                        continue

                # obsługa dodawania
                else:
                    if linia[0] == u"Č":  # przerwa!
                        dodawaj = False
                    elif linia[0] in u"ĚÍ":
                        tabela.append(linia)
                    # koniec tabeli
                    elif linia[0] == u"Đ":
                        tabela.append(linia)
                        jest_tabela = False
                        dodawaj = False
                        continue

            # STOPKA
            else:
                stopka.append(linia)

        return (NaglowekPomiaru(naglowek), TabelaPomiaru(tabela), stopka)


class PomiarDrgania(Pomiar):
    """Reprezentacja Pomiaru II typu uwzględniająca podział na sekcje właściwe temu typowi."""
    def __init__(self, numer, typ, tresc):
        super(PomiarDrgania, self).__init__(numer, typ, tresc)
        # sekcje
        self.naglowek, self.tabela, self.stopka, self.drugi_naglowek, self.druga_tabela, self.druga_stopka = self.wydzielSekcje()

    def wydzielSekcje(self):
        """
        Dzieli treść Pomiaru na sekcje: nagłówek, tabelę, stopkę, drugi naglowek (jeśli jest), drugą tabelę (jeśli jest) oraz drugą stopkę (jeśli jest) i zwraca je w krotce
        """
        naglowek = []
        tabela = []
        stopka = []
        drugi_naglowek = []
        druga_tabela = []
        druga_stopka = []
        # flagi
        jest_naglowek = True
        jest_tabela = False
        dodawaj = False
        poczatek_tab = True  # flaga konieczna dla prawidłowej obsługi linii rozdzielających wiersze tabeli
        jest_drugi_naglowek = False
        jest_druga_tabela = False
        jest_druga_stopka = False

        for linia in self.tresc:
            # NAGŁÓWEK
            if jest_naglowek:
                if linia[0] != u"Č":
                    if jest_drugi_naglowek:
                        drugi_naglowek.append(linia)
                    else:
                        naglowek.append(linia)
                else:
                    jest_naglowek = False
                    jest_tabela = True
                    poczatek_tab = True
                    continue
            # TABELA
            elif jest_tabela:
                # obsługa początku tabeli i końca przerwy
                if not dodawaj:
                    if linia[0] == u"Í":
                        dodawaj = True
                        if poczatek_tab:
                            if jest_druga_tabela:
                                druga_tabela.append(linia)
                            else:
                                tabela.append(linia)
                            poczatek_tab = False
                        continue

                # obsługa dodawania
                else:
                    if linia[0] == u"Č":  # przerwa!
                        dodawaj = False
                    elif linia[0] in u"ĚÍ":
                        if jest_druga_tabela:
                            druga_tabela.append(linia)
                        else:
                            tabela.append(linia)
                    # koniec tabeli
                    elif linia[0] == u"Đ":
                        if jest_druga_tabela:
                            druga_tabela.append(linia)
                        else:
                            tabela.append(linia)
                        jest_tabela = False
                        jest_druga_tabela = False
                        dodawaj = False
                        continue

            # STOPKA (lub powrót do TABELI, jeśli jest jeszcze jedna)
            else:
                if u"Punkt !pomiarowy:" in linia:
                    drugi_naglowek.append(linia)
                    jest_naglowek = True
                    jest_drugi_naglowek = True
                    jest_druga_tabela = True
                    jest_druga_stopka = True
                    continue
                else:
                    if jest_druga_stopka:
                        druga_stopka.append(linia)
                    else:
                        stopka.append(linia)

        if len(drugi_naglowek) == 0:
            drugi_naglowek = None
        if len(druga_tabela) == 0:
            druga_tabela = None
        else:
            druga_tabela = TabelaPomiaru(druga_tabela)
        if len(druga_stopka) == 0:
            druga_stopka = None

        return (NaglowekPomiaru(naglowek), TabelaPomiaru(tabela), stopka, drugi_naglowek, druga_tabela, druga_stopka)


class PomiarPylyChemia(Pomiar):
    """Reprezentacja Pomiaru III typu uwzględniająca podział na sekcje właściwe temu typowi."""
    def __init__(self, numer, typ, tresc):
        super(PomiarPylyChemia, self).__init__(numer, typ, tresc)
        # sekcje
        self.naglowek, self.tabela, self.tabela_ndsch, self.stopka = self.wydzielSekcje()

    def wydzielSekcje(self):
        """
        Dzieli treść Pomiaru na sekcje: nagłówek, tabelę, tabelę NDSCh (jeśli jest) oraz stopkę i zwraca je w krotce
        """
        naglowek = []
        tabela = []
        tabela_ndsch = []
        stopka = []
        # flagi
        jest_naglowek = True
        jest_tabela = False
        jest_ndsch = False
        dodawaj = False
        poczatek_tab = True  # flaga konieczna dla prawidłowej obsługi linii rozdzielających wiersze tabeli

        for linia in self.tresc:
            # NAGŁÓWEK
            if jest_naglowek:
                if linia[0] != u"Č":
                    naglowek.append(linia)
                else:
                    jest_naglowek = False
                    jest_tabela = True
                    continue
            # TABELA
            elif jest_tabela:
                # obsługa początku tabeli i końca przerwy
                if not dodawaj:
                    if linia[0] == u"Í":
                        dodawaj = True
                        if poczatek_tab:
                            if jest_ndsch:
                                tabela_ndsch.append(linia)
                            else:
                                tabela.append(linia)
                            poczatek_tab = False
                        continue

                # obsługa dodawania
                else:
                    if linia[0] == u"Č":  # przerwa!
                        dodawaj = False
                    elif linia[0] in u"ĚÍ":
                        if jest_ndsch:
                            tabela_ndsch.append(linia)
                        else:
                            tabela.append(linia)
                    # koniec tabeli
                    elif linia[0] == u"Đ":
                        if jest_ndsch:
                            tabela_ndsch.append(linia)
                        else:
                            tabela.append(linia)
                        jest_tabela = False
                        jest_ndsch = False
                        dodawaj = False
                        continue

            # STOPKA (lub powrót do TABELI, jeśli NDSch)
            else:
                if u"Ocena !zgodności !stężeń !chwilowych !z !NDSCh:" in linia:
                    jest_tabela = True
                    jest_ndsch = True
                    poczatek_tab = True
                    continue
                else:
                    stopka.append(linia)

        if len(tabela_ndsch) == 0:
            tabela_ndsch = None
        else:
            tabela_ndsch = TabelaPomiaru(tabela_ndsch)

        return (NaglowekPomiaru(naglowek), TabelaPomiaru(tabela), tabela_ndsch, stopka)


class ParserPomiaru(object):
    """
    Klasa zawierająca elementy wspólne dla parsowania wszystkich Pomiarów (parsowanie nagłówka), z której dziedziczą parsery poszczególnych typów Pomiarów
    """
    # stałe klasy
    TEKST_PRZED_DATA = (u"Data !pomiarów...........:", u"Data !pobierania !próbek..........:")

    def __init__(self, pomiar):
        super(ParserPomiaru, self).__init__()
        self.pomiar = pomiar  # Pomiar do sparsowania
        # sparsowane dane wspólne dla wszystkich klas potomnych (zapisane w parserze dla lepszego debuggowania)
        self.nr_spr = self.sparsujNrSprawozdania()
        self.miejsce_pom, self.nazwa_st, self.data_pom = self.sparsujPolaNaglowka()

    # TODO: sprawdzić czy argument dl_wzorca nie jest do wyrzucenia
    @staticmethod
    def scalLinieTekstu(linie_tekstu, wzorce_konca=None):
        """Poprawia i scala (konkatenuje) linie tekstu DRK (najczęściej jest to kolumna wiersza tabeli). Ostatni argument steruje zakończeniem scalania"""
        scalone_linie = []

        for linia in linie_tekstu:
            if len(linia) > 0:
                # usuwamy wykrzykniki
                poprawna_linia = linia.replace(u"!", u"")
                jest_przeniesienie = False

                # obsługa wyjścia z pętli
                if wzorce_konca is not None:
                    if any(poprawna_linia[:len(wzorzec)] == wzorzec for wzorzec in wzorce_konca):
                        break

                # czy jest przeniesienie?
                if len(poprawna_linia) > 0 and poprawna_linia[-1] == u"-":
                    poprawna_linia = poprawna_linia[:-1]
                    jest_przeniesienie = True

                if len(scalone_linie) == 0:
                    scalone_linie.append(poprawna_linia)  # nie potrzebujemy spacji na starcie
                    if jest_przeniesienie:
                        scalone_linie.append(u'Ě')  # znacznik przeniesienia, do usunięcia razem ze spacją przy zamykaniu scalania
                else:
                    scalone_linie.append(u" " + poprawna_linia)
                    if jest_przeniesienie:
                        scalone_linie.append(u'Ě')

        # zamykanie scalania z czyszczeniem znaczników przeniesienia
        return u"".join(scalone_linie).replace(u'Ě ', u"")

    @staticmethod
    def nieMaDanychDoSpars(sekcja_miejsca_i_nazwy):
        """Sprawdza czy w sekcji miejsca i nazwy są dane do sparsowania"""
        nie_ma_danych = False

        if sekcja_miejsca_i_nazwy is not None:
            # obsługa braku danych w drganiach
            for linia in sekcja_miejsca_i_nazwy:
                if u'--------------------' in linia:
                    nie_ma_danych = True
                    break
            # obsługa braku danych w pyłach/chemii
            if not nie_ma_danych and all(not znak.isalnum() for linia in sekcja_miejsca_i_nazwy for znak in linia):
                nie_ma_danych = True
        else:
            nie_ma_danych = True

        return nie_ma_danych

    def sparsujNrSprawozdania(self):
        """Parsuje nr sprawozdania. Zamienia format szczegółowy 000/II/00 na podstawowy 000/00"""
        nr_spr = []  # konkatenacja metodą na listę
        licznik_slashow = 0

        for znak in self.pomiar.tresc[0]:
            # w stringu wynikowym chcemy tylko pierwszego slasha...
            if licznik_slashow == 0 and znak == u'/':
                licznik_slashow += 1
                nr_spr.append(znak)
                continue
            # ...i same liczby
            else:
                if znak.isdigit():
                    nr_spr.append(znak)

        return u"".join(nr_spr)

    def sparsujDatePomiaru(self):
        """Parsuje datę pomiaru"""
        for linia in self.pomiar.naglowek.sekcja_daty:
            if any(znacznik in linia for znacznik in self.TEKST_PRZED_DATA):
                # zwyczajna data
                if u"/" not in linia:
                    linia_z_data = linia[:-2]
                    data_pom = linia_z_data[-12:]
                    break
                else:  # data 2-dniowa
                    data_pom = linia[-15:]  # są 3 znaki więcej do wycięcia i nie ma 2 niepotrzebnych znaków na końcu linii
                    break
        else:
            data_pom = PLACEHOLDERY.data_pom

        return data_pom

    @staticmethod
    def wytnijPoprawnaLinie(linia):
        """
        Do parsowania miejsca pomiaru i nazwy stanowiska (oraz źródła pomiaru i punktu pomiarowego w DRGANIACH). Wycina poprawną linię (zostawia w linii tylko to co nadaje się do parsowania)"""

        # w nagłównku Pomiaru (dokładnie w sekcji miejsca i nazwy) możliwe są 2 rodzaje linii: pełna (wyczerpująca maksymalną ilośc znaków w polu tekstowym HPZ = 72), która kończy się tylko standardowym znacznikiem końca linii '\u2013'
        # oraz krótsza, która kończy się oprócz znacznika jeszcze dodatkowo poprzedzającymi go spacją i znakiem nie będącym małą literą (dla linii o dł. pow. ok. 11 znaków) lub nie będącym dużą literą (dla linii o dł. poniżej ok. 11 znaków).

        # w związku z tym najpierw trzeba sprawdzić z którą linią mamy do czynienia

        do_wyciecia = linia[:-1]  # wszystko bez ostatniego znaku ('\u2013')

        nie_jest_pelna = False
        byl_znak_spec = False
        licznik = 0

        for znak in do_wyciecia[::-1]:  # iteracja od końca!
            # sprawdzamy czy ostatni znak jest dużą lub małą literą (w zależności od długości linii)...
            # UWAGA! 14 - test dotyczy argumentu metody
            if licznik == 0 and len(linia) < 14:
                if not znak.isupper():
                    byl_znak_spec = True
                    licznik += 1
                    continue
            elif licznik == 0 and len(linia) >= 14:
                if not znak.islower():
                    byl_znak_spec = True
                    licznik += 1
                    continue

            # ...jeśli jest znak specjalny i poprzedza go spacja to linia nie jest pełna
            if licznik == 1 and byl_znak_spec and znak == u" ":
                nie_jest_pelna = True
                break
            else:
                break

        # na końcu usunięcie wstawianych przez HPZ wykrzykników
        if nie_jest_pelna:
            return do_wyciecia[:-2]
        else:
            return do_wyciecia

    def sparsujPolaNaglowka(self):
        """
        Parsuje miejsce pomiaru i nazwę stanowiska. Zwraca je oraz datę w formie krotki PolaNaglowka. Wieloliniowe miejsca stają się stringami rozdzielonymi znakiem nowej linii w odpowiednich miejscach, a wieloliniowe nazwy parsowane są do jednej linii
        """

        # ZAŁOŻENIA ALGORYTMU:

        # 1) Test wielkości pierwszej litery. MIEJSCA zawsze zaczynają się wielką literą (jeśli MIEJSCE jest wieloliniowe, to każda nowa linia zaczyna się wielką literą). NAZWY zawsze zaczynają się małą literą. Dla pierwszej parsowanej linii wielka litera zawsze oznacza MIEJSCE a mała NAZWĘ. Jeśli parsowanie zaczyna się od NAZWY, MIEJSCE pozostaje puste. W przypadku kolejnych parsowanych linii mała litera na ogół oznacza, że parsowana linia jest początkiem NAZWY (następuje zamknięcie MIEJSCA i otwarcie NAZWY), ale w niektórych przypadkach może to być zakończenie MIEJSCA (całego bądź tylko jednej z jego linii). O tym czy dana parsowana linia jest: Aa) ostatnim kawałkiem linii MIEJSCA/ostatnim kawałkiem NAZWY lub Ab)początkiem nowej linii MIEJSCA/początkiem NAZWY albo B) kontynuacją otwartej linii MIEJSCA/otwartej NAZWY - decyduje jej długość.

        # 2) Test długości linii. Przyjęty próg: 65 znaków. Zakładamy, że każda parsowana linia krótsza niż próg jest przypadkiem 1A), a każda parsowana linia dłuższa niż próg jest przypadkiem 1B). O tym czy mamy do czynienia z przypadkiem 1Aa) czy 1Ab) decyduje kolejna linia.

        # Przykład:
        # parsowana linia = PL
        # Ewaluacja krok #1: Pierwsza PL zaczyna się wielką literą ===> otwarcie flagi MIEJSCA
        # Ewaluacja krok #2: Jest krótka ===> jest zakończeniem aktualnej (pierwszej) linii MIEJSCA ==> dodanie PL bezpośrednio do MIEJSCA
        # Ewaluacja krok #3: Druga PL zaczyna się wielką literą ===> flaga MIEJSCA pozostaje otwarta
        # Ewaluacja krok #4: Jest długa ===> dodanie PL do aktualnej linii MIEJSCA
        # Ewaluacja krok #5: Trzecia PL zaczyna się małą literą ===> początek NAZWY?
        # Ewaluacja krok #6a: jest długa ===> na pewno początek NAZWY, zamknięcie flagi MIEJSCA, dodanie aktualnej linii MIEJSCA do MIEJSCA, otwarcie flagi NAZWY, dodanie PL do NAZWY
        # Ewaluacja krok #6b: jest krótka ===> początek NAZWY/zakończenie MIEJSCA, zapisanie PL w SCHOWKU
        # Ewaluacja krok #7 (po #6b): czwarta PL zaczyna się małą literą ==> poprzednia PL była zakończeniem MIEJSCA, dodanie SCHOWKA do aktualnej linii MIEJSCA, dodanie aktualnej linii do MIEJSCA, zamknięcie flagi MIEJSCA, otwarcie flagi NAZWY, dodanie PL do nazwy

        # 3) Właściwa konkatenacja linii (z obsługą przeniesień) oddelegowana na koniec do metody scalLinieTekstu()

        miejsce_pom = []  # może mieć więcej niż 1 linię
        aktualna_linia_miejsca = []
        nazwa_st = []  # tylko 1 linia
        schowek = u""
        # flagi
        nazwa_otwarta = False  # otwarta nazwa oznacza jednocześnie zamknięte miejsce

        if not self.nieMaDanychDoSpars(self.pomiar.naglowek.sekcja_miejsca_i_nazwy):
            for x in xrange(len(self.pomiar.naglowek.sekcja_miejsca_i_nazwy)):
                linia = self.pomiar.naglowek.sekcja_miejsca_i_nazwy[x]
                # brane pod uwagę są tylko linie o dł. pow. 3 znaków (najkrótsza linia nadająca się do parsowania zawiera jeden znak + 3-znakowy suffix wstawiony przez HPZ)
                if len(linia) > 3:
                    poprawna_linia = self.wytnijPoprawnaLinie(linia)

                    # test długości linii
                    if any(znak == poprawna_linia[-1] for znak in u":,-") or len(poprawna_linia) > 65:
                        jest_dluga = True
                    else:
                        jest_dluga = False

                    # A) czy linia zaczyna się dużą literą?
                    if poprawna_linia[0].isupper():
                        if nazwa_otwarta:
                            nazwa_st.append(poprawna_linia)
                        else:
                            if jest_dluga:
                                aktualna_linia_miejsca.append(poprawna_linia)
                            else:
                                # linia jest krótka - zamykamy aktualnie zachowaną linię, dodajemy parsowaną
                                if len(aktualna_linia_miejsca) == 0:
                                    # UWAGA na wykrzykniki w tekscie DRK!
                                    # metoda scalLinieTekstu() nie przyjmuje 1 linii a listę linii
                                    lista_do_scalenia = []
                                    lista_do_scalenia.append(poprawna_linia)
                                    miejsce_pom.append(self.scalLinieTekstu(lista_do_scalenia))
                                else:
                                    # flush!
                                    miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
                                    aktualna_linia_miejsca = []

                                    lista_do_scalenia = []
                                    lista_do_scalenia.append(poprawna_linia)
                                    miejsce_pom.append(self.scalLinieTekstu(lista_do_scalenia))

                    # B) czy linia zaczyna się małą literą?
                    elif poprawna_linia[0].islower():

                        # osobny przypadek adresu
                        if poprawna_linia[:4] == u"ul.":
                            aktualna_linia_miejsca.append(poprawna_linia)
                            continue

                        # osobny przypadek kontynuacji przeniesienia MIEJSCA
                        if len(aktualna_linia_miejsca) > 0:
                            ost_element_miejsca = aktualna_linia_miejsca[-1]

                            if ost_element_miejsca[-1] == u"-" and ost_element_miejsca[-2].isalnum():

                                if jest_dluga:
                                    aktualna_linia_miejsca.append(poprawna_linia)
                                    continue
                                else:
                                    aktualna_linia_miejsca.append(poprawna_linia)
                                    # flush!
                                    miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
                                    aktualna_linia_miejsca = []
                                    continue

                        if not nazwa_otwarta:
                            if jest_dluga:
                                # przed dodaniem nazwy zamykamy miejsce
                                if len(aktualna_linia_miejsca) > 0:
                                    # flush!
                                    miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
                                    aktualna_linia_miejsca = []

                                if len(schowek) > 0:
                                    nazwa_st.append(schowek)
                                    schowek = u""

                                nazwa_st.append(poprawna_linia)
                                nazwa_otwarta = True
                            else:
                                if len(schowek) == 0:
                                    schowek = poprawna_linia
                                else:
                                    # przed dodaniem nazwy zamykamy miejsce
                                    aktualna_linia_miejsca.append(schowek)
                                    # flush!
                                    schowek = u""
                                    miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
                                    aktualna_linia_miejsca = []

                                    nazwa_st.append(poprawna_linia)
                                    nazwa_otwarta = True
                        else:
                            nazwa_st.append(poprawna_linia)

                    # C) nie zaczyna się ani małą ani dużą, może zaczyna się na "(", cyfrę lub "- "?
                    else:
                        if nazwa_otwarta:
                            nazwa_st.append(poprawna_linia)
                        else:
                            if jest_dluga:
                                aktualna_linia_miejsca.append(poprawna_linia)
                            else:
                                # linia jest krótka - dodajemy parsowaną i  zamykamy aktualnie zachowaną linię
                                # flush!
                                aktualna_linia_miejsca.append(poprawna_linia)
                                miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
                                aktualna_linia_miejsca = []

            # sprzątanie po zakończeniu pętli
            if len(aktualna_linia_miejsca) > 0:
                miejsce_pom.append(self.scalLinieTekstu(aktualna_linia_miejsca))
            if len(schowek) > 0:
                nazwa_st.append(schowek)

            if len(miejsce_pom) > 0:
                miejsce_pom = u"\n".join(miejsce_pom)
            else:
                miejsce_pom = PLACEHOLDERY.miejsce_pom
                print u"Nieudane parsowanie miejsca pomiaru (Pomiar typu '%s' #%d). Przyjęto wartość domyślną ('%s')" % (self.pomiar.typ, self.pomiar.numer, PLACEHOLDERY.miejsce_pom)

            if len(nazwa_st) > 0:
                nazwa_st = self.scalLinieTekstu(nazwa_st)
            else:
                nazwa_st = PLACEHOLDERY.nazwa_st
                print u"Nieudane parsowanie nazwy stanowiska (Pomiar typu '%s' #%d). Przyjęto wartość domyślną ('%s')" % (self.pomiar.typ, self.pomiar.numer, PLACEHOLDERY.nazwa_st)

        else:
            miejsce_pom = PLACEHOLDERY.miejsce_pom
            nazwa_st = PLACEHOLDERY.nazwa_st

            print u"Nieudane parsowanie miejsca pomiaru i nazwy stanowiska (Pomiar typu '%s' #%d). Przyjęto wartości domyślne ('%s' i '%s')" % (self.pomiar.typ, self.pomiar.numer, PLACEHOLDERY.miejsce_pom, PLACEHOLDERY.nazwa_st)        
        return PolaNaglowkaPom(miejsce_pom, nazwa_st, self.sparsujDatePomiaru())

    @staticmethod
    # metoda używana w klasach potomnych
    def sparsujKrotnosc(wiersz, indeks_rzedu_kol=0):
        """Parsuje i zwraca krotność NDW z właściwej kolumny podanego wiersza tabeli Pomiaru"""
        krotnosc = []
        # krotność znajduje się zawsze w ostatniej kolumnie
        kolumna = wiersz.kolumny[-1]

        for znak in kolumna[indeks_rzedu_kol]:
            if znak in u'<>.-' or znak.isdigit():
                krotnosc.append(znak)

        assert len(krotnosc) > 0, "Nieudane parsowanie krotnosci (Pomiar #%d, stanowisko: %r). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st)

        return u"".join(krotnosc)

    def sparsujCzynnosci(self, tabela):
        """Parsuje opisy czynności z odpowiednich kolumn wierszy podanej tabeli Pomiaru"""
        czynnosci = []  # lista stringów

        # opisy czynności znajdują się w drugiej kolumnie każdego wiersza tabeli Pomiaru
        for wiersz in tabela.wiersze:
            kolumna = wiersz.kolumny[1]

            czynnosci.append(self.scalLinieTekstu(kolumna))

        return czynnosci


class ParserHalasu(ParserPomiaru):
    """Parsuje dane z Pomiaru odpowiedniego typu i tworzy obiekt klasy Halas"""
    # stałe klasy
    TEKST_PRZED_NDP = (u"Poziom !ekspozycji !na !hałas !.........................:", u"Maksymalny !poziom !dźwięku !A !........................:", u"Szczytowy !poziom !dźwięku !C !.........................:")  

    def __init__(self, pomiar):
        super(ParserHalasu, self).__init__(pomiar)

    def sparsujPoziomEkspozycji(self):
        """Parsuje poziom ekspozycji z odpowiedniej pozycji w tabeli Pomiaru"""
        poz_ekspoz = []

        kolumna = self.pomiar.tabela.wiersze[0].kolumny[8]  # przedostatnia kolumna pierwszego wiersza tabeli
        for znak in kolumna[0]:  # poziom ekspozycji jest w pierwszym rzędzie kolumny
            if znak in u'<>.' or znak.isdigit():
                poz_ekspoz.append(znak)

        return u"".join(poz_ekspoz)

    def sparsujSzczytPoziomC(self):
        """Parsuje szczytowy poziom dźwięku C z odpowiednich pozycji w tabeli Pomiaru"""
        wyniki = []

        # szukane dane znajdują się w pierwszym rzędzie ósmej kolumny każdego wiersza tabeli
        for wiersz in self.pomiar.tabela.wiersze:
            kolumna = wiersz.kolumny[7]
            wynik = []
            for znak in kolumna[0]:
                if znak.isdigit():
                    wynik.append(znak)
            if len(wynik) > 0:
                wyniki.append(int(u"".join(wynik)))

        if len(wyniki) > 0:
            posort_wyniki = sorted(wyniki)
            szczyt_C = unicode(posort_wyniki.pop())
        else:
            szczyt_C = PLACEHOLDER_MYSLNIK
            print u"Nieudane parsowanie szczytowego poziomu dźwięku C w HAŁASIE (Pomiar #%d). Przyjęto wartość domyślną: '%s'" % (self.pomiar.numer, szczyt_C)

        return szczyt_C

    def sparsujMaksPoziomA(self):
        """Parsuje maksymalny poziom dźwięku A z odpowiednich pozycji w tabeli Pomiaru"""
        wyniki = []

        # szukane dane znajdują się w pierwszym rzędzie siódmej kolumny każdego wiersza tabeli
        for wiersz in self.pomiar.tabela.wiersze:
            kolumna = wiersz.kolumny[6]
            wynik = []
            for znak in kolumna[0]:
                if znak.isdigit():
                    wynik.append(znak)
            if len(wynik) > 0:
                wyniki.append(int(u"".join(wynik)))

        if len(wyniki) > 0:
            posort_wyniki = sorted(wyniki)
            maks_A = unicode(posort_wyniki.pop())
        else:
            maks_A = PLACEHOLDER_MYSLNIK
            print u"Nieudane parsowanie maksymalnego poziomu dźwięku A w HAŁASIE (Pomiar #%d). Przyjęto wartość domyślną: '%s'" % (self.pomiar.numer, maks_A)

        return maks_A

    @staticmethod
    def filtrujCzynnosci(czynnosci):
        """Usuwa z opisu czynnosci niepotrzebne frazy"""
        przefiltrowane = []
        # frazy = [u" + hałas ogólny", u" + hałas z hali"]
        fraza = u" + hałas"

        # w 99 przypadków na 100 to co znajduje się po 'frazach' jest również niepotrzebne, a więc wystarczy zesplitować string po frazie i zachować tylko lewą część splitu
        for czynnosc in czynnosci:
            if fraza in czynnosc:
                nowa_czynnosc = czynnosc.split(fraza)[0]
                przefiltrowane.append(nowa_czynnosc)
            else:
                przefiltrowane.append(czynnosc)

        przefiltrowane2 = []
        for czynnosc in przefiltrowane:
            if u"przerwy technologiczne" in czynnosc:
                przefiltrowane2.append(czynnosc.replace(u"przerwy technologiczne", u"przerwy"))
            else:
                przefiltrowane2.append(czynnosc)

        return przefiltrowane2

    def sparsujNajwyzszeDopuszczalnePoziomy(self):
        """
        Parsuje najwyższe dopuszczalne poziomy: ekspozycji, maksymalnego dźwięku A oraz szczytowego dźwięku C i oddaje je w krotkce
        """
        ndp = []
        for linia in self.pomiar.stopka:
            if any(znacznik in linia for znacznik in self.TEKST_PRZED_NDP):
                linia_z_ndp = linia[:-6]
                if self.TEKST_PRZED_NDP[0] in linia:
                    ndp_ekspoz = linia_z_ndp[-2:]
                    ndp.append(ndp_ekspoz)
                    continue
                elif self.TEKST_PRZED_NDP[1] in linia:
                    ndp_maks_A = linia_z_ndp[-3:]
                    ndp.append(ndp_maks_A)
                    continue
                else:
                    ndp_szczyt_C = linia_z_ndp[-3:]
                    ndp.append(ndp_szczyt_C)
            if len(ndp) == 3:
                break

        assert len(ndp) == 3, "Nieudane parsowanie najwyzszych dopuszczalnych wartosci w HALASIE (Pomiar #%d, stanowisko: %r). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st)

        return tuple(ndp)

    def podajHalas(self):
        """Tworzy obiekt klasy Halas ze sparsowanych danych i podaje dalej"""
        krotnosc = self.sparsujKrotnosc(self.pomiar.tabela.wiersze[0])  # krotność znajduje się w pierwszym wierszu
        czynnosci = self.filtrujCzynnosci(self.sparsujCzynnosci(self.pomiar.tabela))
        ndp_ekspoz, ndp_maks_A, ndp_szczyt_C = self.sparsujNajwyzszeDopuszczalnePoziomy()
        wykonawca_pom = STALE_ZEWNETRZNE[u"WYKONAWCA_POMIARU"]
        metoda_pom = STALE_ZEWNETRZNE[u"METODY_POMIARÓW"].halas
        poz_ekspoz = self.sparsujPoziomEkspozycji()
        maks_A = self.sparsujMaksPoziomA()
        szczyt_C = self.sparsujSzczytPoziomC()

        halas = Halas(self.nr_spr, self.miejsce_pom, self.nazwa_st, self.data_pom, ndp_ekspoz, poz_ekspoz, krotnosc, czynnosci, wykonawca_pom, metoda_pom, maks_A, szczyt_C, ndp_maks_A, ndp_szczyt_C)

        return halas


# implementacja od wersji 0.5
class ParserDrgan(ParserPomiaru):
    """Parsuje dane z pomiaru odpowiedniego typu"""
    def __init__(self, pomiar):
        super(ParserDrgan, self).__init__(pomiar)

    def wydzielSekcjeZrodla(self):
        """Wydziela sekcję źródła drgań z sekcji daty nagłówka Pomiaru"""
        sekcja_zrodla = []

        dodawaj = False
        for linia in self.pomiar.naglowek.sekcja_daty:
            if u"Źródło !drgań" in linia:
                dodawaj = True
            elif u"Przyjęty !układ !współrzędnych" in linia:
                break
            else:
                if dodawaj:
                    sekcja_zrodla.append(linia)

        return sekcja_zrodla

    def wydzielSekcjePunktuPomiarowego(self):
        """Wydziela sekcję punktu pomiarowego z sekcji daty naglowka Pomiaru"""
        sekcja_punktu_pom = []

        dodawaj = False
        for linia in self.pomiar.naglowek.sekcja_daty:
            if u"Punkt !pomiarowy" in linia:
                dodawaj = True
            elif u"------------------" in linia:
                break
            else:
                if dodawaj:
                    sekcja_punktu_pom.append(linia)

        return sekcja_punktu_pom

    def sparsujZrodlo(self):
        """Parsuje i zwraca źródło (lub źródła) pomiaru"""
        zrodlo_pom = []

        for linia in self.wydzielSekcjeZrodla():
            poprawna_linia = self.wytnijPoprawnaLinie(linia).replace(u"!", u"")
            zrodlo_pom.append(poprawna_linia)

        return u"\n".join(zrodlo_pom)

    def sparsujPunktPomiarowy(self):
        """Parsuje i zwraca punkt pomiarowy"""
        punkt_pom = []

        for linia in self.wydzielSekcjePunktuPomiarowego():
            poprawna_linia = self.wytnijPoprawnaLinie(linia).replace(u"!", u"")
            punkt_pom.append(poprawna_linia)

        return u"\n".join(punkt_pom)

    # nieskończone
    def sparsujEkspozycje(self, stopka):
        """Parsuje i zwraca ekspozycję (dzienną lub 30-minutową) z podanej stopki Pomiaru"""
        # ekspozycja = []

        # for linia in stopka:
        #     if u"Ekspozycja" in linia and u"......" in linia:


# TODO: stężenia powyżej oznaczalności
class ParserPylowChemii(ParserPomiaru):
    """Parsuje dane z Pomiaru odpowiedniego typu i tworzy obiekt klasy Pylochem"""
    # stałe klasy
    INDEKS_RZEDU_RESP = 4

    def __init__(self, pomiar):
        super(ParserPylowChemii, self).__init__(pomiar)
        self.nazwy_czynnikow, self.nazwy_czynnikow_ndsch = self.ustalNazwyCzynnikow()
        self.czynnosci = self.sparsujCzynnosci()
        self.ilosc_wystapien_p_o = self.ustalIloscWystapienPO()

    def sparsujCzynnosci(self):
        """Parsuje i zwraca opisy czynnosci z właściwej kolumny podanego wiersza"""
        wzorce = [u"Czas ekspozycji [min]"]
        # opisy czynności znajdują się w drugiej kolumnie pierwszego wiersza tabeli Pomiaru
        kolumna = self.pomiar.tabela.wiersze[0].kolumny[1]
        czynnosci = self.scalLinieTekstu(kolumna, wzorce)
        return czynnosci

    def sparsujNazweCzynnika(self, wiersz):
        """Parsuje i zwraca nazwę Czynnika z właściwej kolumny podanego wiersza"""
        wzorce = [u"Zaw. ", u"-met.", u"NDSch"]
        # nazwa znajduje się w pierwszej kolumnie
        kolumna = wiersz.kolumny[0]
        nazwa_czynnika = self.scalLinieTekstu(kolumna, wzorce)
        # usunięcie "*" z nazwy
        return nazwa_czynnika.replace(u"*", u"")

    def ustalNazwyCzynnikow(self):
        """
        Ustala nazwy Czynników dla parsowanego Pomiaru i zapisuje je jako krotkę (1 listę i obiekt None jeśli Pomiar posiada tylko tabelę lub 2 listy jeśli Pomiar posiada również tabelę NDSch) w tym parserze
        """
        nazwy_czynnikow = []
        nazwy_czynnikow_ndsch = []

        for wiersz in self.pomiar.tabela.wiersze:
            nazwy_czynnikow.append(self.sparsujNazweCzynnika(wiersz))

        if self.pomiar.tabela_ndsch is None:
            nazwy_czynnikow_ndsch = None
        else:
            for wiersz in self.pomiar.tabela_ndsch.wiersze:
                nazwy_czynnikow_ndsch.append(self.sparsujNazweCzynnika(wiersz))

        return (nazwy_czynnikow, nazwy_czynnikow_ndsch)

    def sparsujNDS(self, wiersz, wzorce=WzorceParsowaniaNDS.ZWYKLY):
        """Parsuje i zwraca NDS wg podanego wzorca z podanej kolumny wiersza"""
        nds = []
        # NDS znajdujw się w pierwszej kolumnie
        kolumna = wiersz.kolumny[0]
        dodawaj = False

        for rzad in kolumna:
            poprawny_rzad = rzad.replace(u"!", u"")
            if any(wzorzec in poprawny_rzad for wzorzec in wzorce):
                # obsługa braku NDS dla pyłu respirabilnego
                if u"brak" in poprawny_rzad:
                    nds = u"brak"
                    break
                else:
                    for znak in poprawny_rzad:
                        if znak == u":":
                            dodawaj = True
                            continue
                        if dodawaj:
                            if znak == u"." or znak.isdigit():
                                nds.append(znak)
                            else:
                                if not znak.isspace():
                                    break
        if len(nds) == 0:
            # obsługa pyłów z nieoznaczonym SiO2, w przypadku których nie ma w tabeli podanego NDS
            nds = PLACEHOLDER_MYSLNIK
            print u"Nieudane parsowanie wartości NDS w PYŁACH/CHEMII (Pomiar #%d). Przyjęto wartość domyślną: '%s'" % (self.pomiar.numer, nds)
            return nds
        elif nds == u"brak":
            return nds
        else:
            return u"".join(nds)

    def sparsujWskaznikNarazenia(self, wiersz):
        """
        Parsuje i zwraca wskaźnik narażenia z właściwej kolumny podanego wiersza. Parsowany jest pełen string wskaźnika narażenia (całość tego, co pojawia się w tabelce karty) - niezależnie od metody poboru ('Cw' dla dozymetrycznej, 'Xgw', 'Dgw' i 'Ggw' dla stacjonarnej)
        """
        cw = []
        xgw = []
        dgw = []
        ggw = []
        # wskaźnik narażenia znajduje się zawsze w przedostatniej kolumnie
        kolumna = wiersz.kolumny[-2]

        # Cw/Xgw jest w pierwszym rzędzie kolumny
        for znak in kolumna[0]:
            if znak in u'<>.=CwXg' or znak.isdigit():
                cw.append(znak)
        # Cw czy Xgw?
        if len(cw) > 0:
            if cw[0] == u"C":
                xgw = None
            else:
                xgw = cw
                cw = None
        else:
            cw = None
            xgw = None

        assert cw is not None or xgw is not None, "Nieudane parsowanie wskaznika narazenia w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r, wiersz tabeli #%d). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st, wiersz.numer)

        # DGw jest w trzecim a GGw w czwartym rzędzie kolumny (o ile w ogóle jest czwarty rząd! (trzy to minimum) [poprawka wrzesień 2015])
        rzedy_do_spars = kolumna[2:4]
        if len(rzedy_do_spars) > 1:
            for i in xrange(2):
                rzad = rzedy_do_spars[i]
                for znak in rzad:
                    if znak in u'<>.=DGw' or znak.isdigit():
                        if i == 0:
                            dgw.append(znak)
                        else:
                            ggw.append(znak)

        if len(dgw) == 0:
            dgw = None
        if len(ggw) == 0:
            ggw = None

        if all(wskaznik is not None for wskaznik in (xgw, dgw, ggw)):
            return u"".join(xgw) + u"; " + u"".join(dgw) + u"; " + u"".join(ggw)
        else:
            if cw is not None:
                return u"".join(cw)
            else:
                return u"".join(xgw)

    def sprawdzCzyNowyPyl(self, wiersz):
        """Sprawdza czy parsowany Czynnik jest nowym pyłem"""
        jest_nowym = False
        kolumna = wiersz.kolumny[0]  # szukana informacja jest pierwszej kolumnie

        for rzad in kolumna:
            poprawny_rzad = rzad.replace(u"!", u"")
            if u"(wdch.)" in poprawny_rzad:
                jest_nowym = True
                break

        return jest_nowym

    # TODO: przetestować
    @staticmethod
    def ustalWzorzecPO(nazwa_czynnika):
        """Ustala wzorzec parsowania p.o. na podstawie danej nazwy Czynnika"""
        znaczniki = (u" i jego", u" i jej", u" – ", u" - ", u" (", u" metaliczny")

        mozliwe_wzorce = []
        for i in xrange(len(znaczniki)):
            znacznik = znaczniki[i]
            if znacznik in nazwa_czynnika:
                mozliwy_wzorzec = nazwa_czynnika.split(znacznik, 1)[0]
                # jeśli delimiter splitu wystąpił w stringu później niż powinien, uzyskany wzorzec będzie składał się ze zbyt wielu wyrazów, a więc...
                # ... przyjmijmy próg 3
                if len(mozliwy_wzorzec.split()) > 3:
                    continue
                else:
                    mozliwe_wzorce.append(mozliwy_wzorzec)
                    continue

        if len(mozliwe_wzorce) == 0:
            wzorzec = nazwa_czynnika
        elif len(mozliwe_wzorce) == 1:
            wzorzec = mozliwe_wzorce[0]
        else:
            # wybieramy najkrótszy z wyciętych wzorców
            posort_mozliwe_wzorce = sorted(mozliwe_wzorce, key=len)
            wzorzec = posort_mozliwe_wzorce[0]

        return wzorzec

    def ustalIloscWystapienPO(self):
        """
        Sprawdza ilość wystąpień informacji o p.o. w stopce Pomiaru (od czego zależy parsowanie p.o)
        """
        ilosc_wystapien = 0
        for linia in self.pomiar.stopka:
            if u"poniżej" in linia:
                ilosc_wystapien += 1

        return ilosc_wystapien

    # TODO: sprawdzić obsługa nowych pyłów jest prawidłowa (czy linia w DRK rzeczywiście wygląda tak: "- dla pyłów - frakcja wdychalna/respirabilna")
    # TODO: ponieważ TT leci w kulki i robi dopiski "- frakcja wdychalna/respirabilna" tylko w przypadkach gdy są oba p.o. (w przypadku gdy jest jeden pozostawia samą nazwę substancji (np. "dla manganu")) - trzeba dodać logikę obsługi tych sytuacji - parser musi sprawdzić ile razu występuje w stopce słowo kluczowe i dodawać suffix tylko jeśli wystąpi więcej niż raz
    # TODO: sprawdzić parsowanie p_o_chw (na przykładzie 15-039-2)
    def sparsujPO(self, wzorzec=None, tryb=None):
        """
        Parsuje i zwraca wartość p.o. ze stopki. Jeśli trzeba, wyznacza linie do sparsowania na podstawie podanego wzorca i trybu parsowania. Parsuje wszystkie możliwe wartości p.o. - również p.o. pyłu respirabilnego oraz próbek chwilowych - o tym jakie p.o. jest zwracane decydują podany wzorzec i tryb
        """

        # funkcja podręczna (tylko dla przejrzystości kodu)
        def sparsujLinie(linia):
            p_o_dozwrotu = []
            # dzielimy linię na dwie części...
            porcje = poprawna_linia.split(u"mg/m")
            # ...i parsujemy lewą od prawej strony
            odwr_porcja = porcje[0][::-1]

            for znak in odwr_porcja:
                # wyjście z pętli
                if znak == u"j":
                    break
                if znak == u"." or znak.isdigit():
                    p_o_dozwrotu.append(znak)

            return p_o_dozwrotu

        p_o = []
        # UWAGA: zastosowany algorytm nie obejmuje przypadków, w których wzorzec (nazwa danego związku) jest na tyle długi (powyżej 32 znaków), że przechodzi w pliku DRK do kolejnej linii. Jest to dość istotny problem, bo wzorzec wraz z dodanym suffiksem (w przypadku frakcji wsychalnych, respirabilnych i chwilówek) na ogół przekroczy ten limit. TODO: sprawdzić precyzyjnie w HPZ-ecie tę ilość znaków.
        # Dodatkowym problemem jest obsługa sytuacji w których powinny wystąpić 2 suffiksy jednocześnie (np. dla tlenków żelaza - frakcja wdychalna oraz próbka chwilowa) - z jednej strony łamie to zastosowany algorytm, z drugiej oznacza prawie pewne przekroczenie limitu 32 znaków w linii
        if wzorzec is not None:
            slownik_odmian = STALE_ZEWNETRZNE[u"SŁOWNIK_ODMIAN"]
            odm_wzorzec = slownik_odmian[wzorzec.lower()]

            # dodanie spacji na początku eliminuje fałszywe trafienia dla nazw Czynników o tym samym rdzeniu a różnych przedrostkach - np. 'tlenek i ditlenek azotu'
            odm_wzorzec = u" " + odm_wzorzec

            # tej pętli nie da się uprościć (eliminując powtórzenia, ponieważ zachowanie gałęzi 'else' jest inne od pozostałych trzech)
            for linia in self.pomiar.stopka:
                poprawna_linia = linia.replace(u"!", u"")

                # obsługa trybu frakcji wdychalnych
                if tryb == TrybParsowaniaPO.FRAKCJA_WDYCH:
                    odm_wzorzec_suffix = odm_wzorzec + u" - frakcja w"
                    if odm_wzorzec_suffix in poprawna_linia and len(odm_wzorzec_suffix) <= 32:
                        p_o = sparsujLinie(poprawna_linia)
                        break
                    else:
                        p_o.append(PLACEHOLDER_MYSLNIK)
                        break

                # obsługa trybu frakcji respirabilnych
                elif tryb == TrybParsowaniaPO.FRAKCJA_RESP:
                    odm_wzorzec_suffix = odm_wzorzec + u" - frakcja r"
                    if odm_wzorzec_suffix in poprawna_linia and len(odm_wzorzec_suffix) <= 32:
                        p_o = sparsujLinie(poprawna_linia)
                        break
                    else:
                        p_o.append(PLACEHOLDER_MYSLNIK)
                        break

                # obsługa trybu chwilówek
                elif tryb == TrybParsowaniaPO.CHWILOWKI:
                    odm_wzorzec_suffix = odm_wzorzec + u" - pr"
                    if odm_wzorzec_suffix in poprawna_linia and len(odm_wzorzec_suffix) <= 32:
                        p_o = sparsujLinie(poprawna_linia)
                        break
                    else:
                        p_o.append(PLACEHOLDER_MYSLNIK)
                        break

                # obsługa trybu typowego
                else:
                    if odm_wzorzec in poprawna_linia and len(odm_wzorzec) <= 32:
                        # ominięcie chwilówek
                        if u" - pr" in poprawna_linia:
                            continue
                        else:
                            p_o = sparsujLinie(poprawna_linia)
                            break

            if len(p_o) == 0:
                print u"Nieudane parsowanie p.o. w PYŁACH/CHEMII (Pomiar #%d). Przyjęto wartość domyślną: '%s'" % (self.pomiar.numer, PLACEHOLDER_MYSLNIK)
                p_o.append(PLACEHOLDER_MYSLNIK)

        # jeśli p.o. pojawiło się w całym Pomiarze tylko raz, mamy wtedy prosty format informacji o p.o., w którym nie potrzeba wzorca
        else:
            for linia in self.pomiar.stopka:
                if u"mg/m" in linia:
                    poprawna_linia = linia.replace(u"!", u"")
                    p_o = sparsujLinie(poprawna_linia)
                    break

            if len(p_o) == 0:
                print u"Nieudane parsowanie p.o. w PYŁACH/CHEMII (pomiar #%d). Przyjęto wartość domyślną: '%s'" % (self.pomiar.numer, PLACEHOLDER_MYSLNIK)
                p_o.append(PLACEHOLDER_MYSLNIK)
        # zwracamy przywracając prawidłową kolejność znaków (patrz: funkcja sparsujLinie)
        return u"".join(p_o[::-1])

    def podajPylochem(self, wiersz):
        """
        Tworzy Czynnik odpowiedniego typu z danych sparsowanych przez rodzica oraz z podanego iersza tabeli Pomiaru i podaje dalej
        """
        jestem_pylem = False
        nazwa_czynnika = self.nazwy_czynnikow[wiersz.numer - 1]  # numer wiersza nie jest indeksem wiersza
        krotnosc = self.sparsujKrotnosc(wiersz)
        nds = self.sparsujNDS(wiersz)
        wykonawca_pom = STALE_ZEWNETRZNE[u"WYKONAWCA_POMIARU"]

        # wybór metody Pomiaru
        if u"pyły" in nazwa_czynnika.lower():
            metoda_pom = STALE_ZEWNETRZNE[u"METODY_POMIARÓW"].pyl
            jestem_pylem = True
        else:
            metoda_pom = STALE_ZEWNETRZNE[u"METODY_POMIARÓW"].chemia

        numer_czynnika = wiersz.numer
        wskaznik_nar = self.sparsujWskaznikNarazenia(wiersz)

        # p.o.
        if self.ilosc_wystapien_p_o > 0:
            # dodatkowy check - jeśli jest p.o., wskaźnik narażenia będzie zawierał znak mniejszości
            if u"<" in wskaznik_nar:
                if self.ilosc_wystapien_p_o == 1:
                    p_o = self.sparsujPO()  # proste parsowanie p.o.
                else:
                    # skomplikowane parsowanie p.o. - trzeba wydobyć ze stopki wartość p.o. z tekstu o nienajwygodniejszym formatowaniu
                    if jestem_pylem:
                        if self.sprawdzCzyNowyPyl(wiersz):
                            p_o = self.sparsujPO(u"PC_N")
                        else:
                            p_o = self.sparsujPO(u"PC")
                    else:
                        wzorzec_p_o = self.ustalWzorzecPO(nazwa_czynnika)
                        # UWAGA! ponieważ TT nie stosuje tej zasady równie konsekwentnie co ja tutaj, muszę ograniczyć stosowanie trybów WDYCH i RESP tylko do manganu i tritlenku glinu (Czynników które mają NDS zarówno dla frakcji wdychalnej jak i respirabilnej). Przynajmniej do czasu znalezienia innego rozwiązania
                        filtr = (u"mangan", u"tritlenek glinu")
                        if u"wdych" in nazwa_czynnika and any(znacznik in nazwa_czynnika.lower() for znacznik in filtr):
                            p_o = self.sparsujPO(wzorzec_p_o, TrybParsowaniaPO.FRAKCJA_WDYCH)
                        elif u"resp" in nazwa_czynnika and any(znacznik in nazwa_czynnika.lower() for znacznik in filtr):
                            p_o = self.sparsujPO(wzorzec_p_o, TrybParsowaniaPO.FRAKCJA_RESP)
                        else:
                            p_o = self.sparsujPO(wzorzec_p_o)
            else:
                p_o = None
        else:
            p_o = None

        pylochem = Pylochem(nazwa_czynnika, self.nr_spr, self.miejsce_pom, self.nazwa_st, self.data_pom, nds, wskaznik_nar, krotnosc, self.czynnosci, wykonawca_pom, metoda_pom, numer_czynnika, p_o)

        return pylochem

    # METODY DLA CZYNNIKÓW KLASY PYŁSIO2
    def sparsujZawartoscSiO2(self, wiersz):
        """Parsuje i zwraca zawartość SiO2 z właściwej kolumny podanego wiersza"""
        zaw_SiO2 = []
        # Zawartość SiO2 znajdujw się w pierwszej kolumnie
        kolumna = wiersz.kolumny[0]
        dodawaj = False
        bylo_Si02 = False

        for rzad in kolumna:
            poprawny_rzad = rzad.replace(u"!", u"")
            if u"Zaw. SiO" in poprawny_rzad:
                bylo_Si02 = True
                if u"nie ozn" in poprawny_rzad:
                    zaw_SiO2 = u"nie oznaczono"
                    break
                else:
                    for znak in poprawny_rzad:
                        if znak == u":":
                            dodawaj = True
                            continue
                        if dodawaj:
                            if znak == u"." or znak.isdigit():
                                zaw_SiO2.append(znak)
                            else:
                                if not znak.isspace():
                                    break

        if bylo_Si02:
            assert len(zaw_SiO2) > 0 or zaw_SiO2 == u"nie oznaczono", "Nieudane parsowanie zawartosci SiO2 w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r, wiersz tabeli #%d). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st, wiersz.numer)

            return u"".join(zaw_SiO2)
        else:
            # zwracamy obiekt None w sytuacji gdy w całej parsowanej kolumnie nie było frazy 'Zaw. SiO' - jest to potrzebne przy parsowaniu danego wiersza tabeli do decyzji czy trzeba parsować Czynnik typu 'PylSiO2' czy nie
            return None

    def sparsujWskaznikNarazeniaResp(self, wiersz):
        """
        Parsuje i zwraca wskaźnik narażenia pyłu respirabilnego z właściwej kolumny podanego wiersza
        """
        cw = []
        # wskaźnik narażenia znajduje się zawsze w przedostatniej kolumnie
        kolumna = wiersz.kolumny[-2]

        # Cw jest w piątym rzędzie kolumny
        for znak in kolumna[self.INDEKS_RZEDU_RESP]:
            if znak in u'<>.=Cw' or znak.isdigit():
                cw.append(znak)

        assert len(cw) > 0, "Nieudane parsowanie wskaznika narazenia pylu respirabilnego w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r, wiersz tabeli #%d). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st, wiersz.numer)

        return u"".join(cw)

    # podajemy zmienną 'zaw_SiO2' jako jeden z argumentów, ponieważ trzeba ją sparsować przed wywołaniem poniższej metody (dla ustalenia czy mamy ją w ogóle wywołać)
    def podajPylSiO2(self, wiersz, zaw_SiO2):
        """
        Tworzy Czynnik odpowiedniego typu z danych sparsowanych przez rodzica oraz z podanego wiersza tabeli Pomiaru i podaje dalej
        """
        nazwa_czynnika = self.nazwy_czynnikow[wiersz.numer - 1]  # numer wiersza nie jest indeksem wiersza
        krotnosc = self.sparsujKrotnosc(wiersz)
        nds = self.sparsujNDS(wiersz)
        wykonawca_pom = STALE_ZEWNETRZNE[u"WYKONAWCA_POMIARU"]
        metoda_pom = STALE_ZEWNETRZNE[u"METODY_POMIARÓW"].pyl_SiO2
        numer_czynnika = wiersz.numer
        wskaznik_nar = self.sparsujWskaznikNarazenia(wiersz)

        # pola wyróżniające klasę PylSiO2
        nds_resp = self.sparsujNDS(wiersz, WzorceParsowaniaNDS.RESP)
        wskaznik_nar_resp = self.sparsujWskaznikNarazeniaResp(wiersz)
        krotnosc_resp = self.sparsujKrotnosc(wiersz, self.INDEKS_RZEDU_RESP)

        # p_o i p_o_resp
        if self.ilosc_wystapien_p_o > 0:
            if u"<" in wskaznik_nar:
                if self.ilosc_wystapien_p_o == 1:
                    p_o = self.sparsujPO()
                else:
                    if self.sprawdzCzyNowyPyl(wiersz):
                        p_o = self.sparsujPO(u"PC_N")
                    else:
                        p_o = self.sparsujPO(u"PC")
            else:
                p_o = None

            if u"<" in wskaznik_nar_resp:
                if self.ilosc_wystapien_p_o == 1:
                    p_o_resp = self.sparsujPO()
                else:
                    if self.sprawdzCzyNowyPyl(wiersz):
                        p_o_resp = self.sparsujPO(u"PR_N")
                    else:
                        p_o_resp = self.sparsujPO(u"PR")

            else:
                p_o_resp = None
        else:
            p_o = None
            p_o_resp = None

        pyl_SiO2 = PylSiO2(nazwa_czynnika, self.nr_spr, self.miejsce_pom, self.nazwa_st, self.data_pom, nds, wskaznik_nar, krotnosc, self.czynnosci, wykonawca_pom, metoda_pom, numer_czynnika, p_o, zaw_SiO2, nds_resp, wskaznik_nar_resp, krotnosc_resp, p_o_resp)

        return pyl_SiO2

    # METODY DLA CZYNNIKÓW KLASY CHEMIACHWILOWKI
    def sparsujWskaznikNarazeniaChw(self, wiersz):
        """
        Parsuje i zwraca wskaźnik narażenia próbki chwilowej z właściwej kolumny podanego wiersza"""
        wyniki = []
        # szukane dane znajdują się w przedostatniej kolumnie tabeli NDSch Pomiaru
        kolumna = wiersz.kolumny[-2]

        for rzad in kolumna:
            if u"Cch" in rzad:
                wynik = []
                for znak in rzad:
                    if znak in u"<>." or znak.isdigit():
                        wynik.append(znak)

                wynik = u"".join(wynik)
                if u"<" in wynik or u">" in wynik:
                    wyniki.append(wynik)
                    break  # nie będzie drugiego wyniku - wychodzimy
                    # TODO: dlaczego w ogóle miałby być drugi wynik??? Przecież wskaźnik jest zawsze tylko 1 na wiersz [wrzesień 2015]
                else:
                    wyniki.append(float(wynik))

        assert len(wyniki) > 0, "Nieudane parsowanie wskaznika narażenia próbki chwilowej w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r, wiersz tabeli NDSch #%d). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st, wiersz.numer)

        if len(wyniki) >= 1:
            posort_wyniki = sorted(wyniki)
            wskaznik_nar_chw = unicode(posort_wyniki.pop())
            # poprawienie formatu stringa na 'X.XX' (dodanie brakującego zera na końcu)
            if len(wskaznik_nar_chw.split(u".")[1]) == 1:
                wskaznik_nar_chw = wskaznik_nar_chw + u"0"
        # TODO: 2 wiersze poniżej po sprawdzeniu czy wszystko działa - do usunięcia
        # elif len(wyniki) == 1:
        #     wskaznik_nar_chw = wyniki[0]

        return wskaznik_nar_chw

    # TODO: coś tu jest ewidentnie pokopane. Dlaczego nie parsuje tego ogolna metoda sparsujKrotnosc() ??? [wrzesień 2015]
    def sparsujKrotnoscChw(self, wiersz):
        """
        Parsuje i zwraca krotność wartości dopuszczalnej dla próbki chwilowej z właściwej kolumny podanego wiersza
        """
        wyniki = []
        # szukane dane znajdują się w ostatniej kolumnie tabeli NDSch Pomiaru
        kolumna = wiersz.kolumny[-1]

        for rzad in kolumna:
            wynik = []
            for znak in rzad:
                if znak in u"<>." or znak.isdigit():
                    wynik.append(znak)

            if len(wynik) > 0:
                wynik = u"".join(wynik)
                if u"<" in wynik or u">" in wynik:
                    wyniki.append(wynik)
                    break  # nie będzie drugiego wyniku - wychodzimy
                    # TODO: dlaczego w ogóle miałby być drugi wynik??? Przecież krotność jest zawsze tylko 1 na wiersz [wrzesień 2015]
                else:
                    wyniki.append(float(wynik))

        assert len(wyniki) > 0, "Nieudane parsowanie krotnosci próbki chwilowej w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r, wiersz tabeli NDSch #%d). Sprawdz czy plik DRK ma odpowiedni format." % (self.pomiar.numer, self.nazwa_st, wiersz.numer)

        if len(wyniki) >= 1:
            posort_wyniki = sorted(wyniki)
            krotnosc_chw = unicode(posort_wyniki.pop())
            # poprawienie formatu stringa na 'X.XX' (dodanie brakującego zera na końcu)
            if len(krotnosc_chw.split(u".")[1]) == 1:
                krotnosc_chw = krotnosc_chw + u"0"
        # TODO: 2 wiersze poniżej po sprawdzeniu czy wszystko działa - do usunięcia
        # elif len(wyniki) == 1:
        #     krotnosc_chw = wyniki[0]

        return krotnosc_chw

    def podajChemieChwilowkowa(self, wiersz, wiersz_ndsch):
        """
        Tworzy Czynnik odpowiedniego typu z danych sparsowanych przez rodzica oraz z podanych wierszy (tabeli i tabeli NDSch) Pomiaru i podaje dalej
        """
        nazwa_czynnika = self.nazwy_czynnikow[wiersz.numer - 1]  # numer wiersza nie jest indeksem wiersza
        krotnosc = self.sparsujKrotnosc(wiersz)
        nds = self.sparsujNDS(wiersz)
        wykonawca_pom = STALE_ZEWNETRZNE[u"WYKONAWCA_POMIARU"]
        metoda_pom = STALE_ZEWNETRZNE[u"METODY_POMIARÓW"].chemia
        numer_czynnika = wiersz.numer
        wskaznik_nar = self.sparsujWskaznikNarazenia(wiersz)

        # pola wyróżniające klasę ChemiaChwilowki
        nds_chw = self.sparsujNDS(wiersz_ndsch, WzorceParsowaniaNDS.CHW)
        wskaznik_nar_chw = self.sparsujWskaznikNarazeniaChw(wiersz_ndsch)
        krotnosc_chw = self.sparsujKrotnoscChw(wiersz_ndsch)

        # p_o i p_o_chw
        if self.ilosc_wystapien_p_o > 0:
            if u"<" in wskaznik_nar:
                if self.ilosc_wystapien_p_o == 1:
                    p_o = self.sparsujPO()
                else:
                    wzorzec_p_o = self.ustalWzorzecPO(nazwa_czynnika)
                    # if u"wdych" in nazwa_czynnika:
                    #     p_o = self.sparsujPO(wzorzec_p_o, TrybParsowaniaPO.FRAKCJA_WDYCH)
                    # elif u"resp" in nazwa_czynnika:
                    #     p_o = self.sparsujPO(wzorzec_p_o, TrybParsowaniaPO.FRAKCJA_RESP)
                    # else:
                    p_o = self.sparsujPO(wzorzec_p_o)
            else:
                p_o = None

            if u"<" in wskaznik_nar_chw:
                if self.ilosc_wystapien_p_o == 1:
                    p_o_chw = self.sparsujPO()
                else:
                    wzorzec_p_o_chw = self.ustalWzorzecPO(nazwa_czynnika)
                    p_o_chw = self.sparsujPO(wzorzec_p_o_chw, TrybParsowaniaPO.CHWILOWKI)
            else:
                p_o_chw = None
        else:
            p_o = None
            p_o_chw = None

        chemia_chwilowkowa = ChemiaChwilowkowa(nazwa_czynnika, self.nr_spr, self.miejsce_pom, self.nazwa_st, self.data_pom, nds, wskaznik_nar, krotnosc, self.czynnosci, wykonawca_pom, metoda_pom, numer_czynnika, p_o, nds_chw, wskaznik_nar_chw, krotnosc_chw, p_o_chw)

        return chemia_chwilowkowa

    def podajCzynniki(self):
        """
        Parsuje odpowiednie Czynniki z każdego wiersza Pomiaru. Zwraca listę sparsowanych Czynników"""
        czynniki = []

        # czy w ogóle są Czynniki NDSch?
        if self.nazwy_czynnikow_ndsch is None:
            for wiersz in self.pomiar.tabela.wiersze:
                # czy to PiSO2?
                zaw_SiO2 = self.sparsujZawartoscSiO2(wiersz)
                if zaw_SiO2 is not None:
                    czynniki.append(self.podajPylSiO2(wiersz, zaw_SiO2))
                else:
                    czynniki.append(self.podajPylochem(wiersz))
        else:
            for wiersz in self.pomiar.tabela.wiersze:
                # czy to PiSO2?
                zaw_SiO2 = self.sparsujZawartoscSiO2(wiersz)
                if zaw_SiO2 is not None:
                    czynniki.append(self.podajPylSiO2(wiersz, zaw_SiO2))
                else:
                    # czy to ChemiaChilowka? Sprawdzanie po nazwach
                    nazwa_czynnika = self.nazwy_czynnikow[wiersz.numer - 1]
                    indeks_wiersza_ndsch = -1

                    for i in xrange(len(self.nazwy_czynnikow_ndsch)):
                        nazwa_czynnika_ndsch = self.nazwy_czynnikow_ndsch[i]

                        if nazwa_czynnika_ndsch == nazwa_czynnika:
                            indeks_wiersza_ndsch = i
                            break

                    if indeks_wiersza_ndsch == -1:
                        # to nie jest ChemiaChwilowka
                        czynniki.append(self.podajPylochem(wiersz))
                    else:
                        wiersz_ndsch = self.pomiar.tabela_ndsch.wiersze[indeks_wiersza_ndsch]
                        czynniki.append(self.podajChemieChwilowkowa(wiersz, wiersz_ndsch))

        assert len(czynniki) > 0, "Nieudane parsowanie listy Czynników w PYLACH/CHEMII (Pomiar #%d, stanowisko: %r). Trudno powiedziec, co sie stalo" % (self.pomiar.numer, self.nazwa_st)

        return czynniki