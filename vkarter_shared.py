# -*- coding: utf-8 -*-
from collections import namedtuple
import codecs

# klasy pomocnicze
GlTypyPomiarow = namedtuple('GlTypyPomiarow', 'halas drgania pyly_chemia')
TypyPomiarow = namedtuple('TypyPomiarow', 'halas drgania pyly chemia pyly_i_chemia')
PodtypyPylochem = namedtuple('PodtypyPylochem', 'pyly chemia pyly_i_chemia')  # na razie nieużywane
PolaNaglowkaPom = namedtuple('PolaNaglowkaPom', 'miejsce_pom nazwa_st data_pom')
MetodyPomiarow = namedtuple('MetodyPomiarow', 'halas pyl pyl_SiO2 chemia')

# stałe parsowania
ZNACZNIKI_POM = TypyPomiarow(u'** !HAŁAS !- !BADANIE !"A" !**', u'** !DRGANIA !- !BADANIE !"A" !**', u'** !PYŁY !**', u'** !SUBSTANCJE !CHEMICZNE !**', u'** !PYŁY !I !SUBSTANCJE !CHEMICZNE !**')

PLACEHOLDERY = PolaNaglowkaPom(u"[miejsce_pomiaru]", u"[nazwa_stanowiska]", u"[data_pomiaru]")

PLACEHOLDER_MYSLNIK = u"-"


class TypPomiaru(object):
    """Enum określający typ Pomiaru lub brak Pomiaru"""
    ZADEN = 0
    HALAS = 1
    DRGANIA = 2
    PYLY_CHEMIA = 3


class WzorceParsowaniaNDS(object):
    """
    Enum określający typ wzorców do parsowania NDS. Wykorzystywany przez metodę sparsujNDS() klasy ParserPylowChemii
    """
    # potrzebujemy list jednoelementowych ze wzgl. na strukturę algorytmu parsowania
    ZWYKLY = [u"NDS :", u"NDS(cał", u"NDS(wdch"]
    RESP = [u"NDS(resp"]
    CHW = [u"NDSch :"]


class TrybParsowaniaPO(object):
    """
    Enum określający tryb parsowania p.o.: 1) frakcja wdychalna, 2) frakcja respirabilna, 3) próbka chwilowa - dla metody sparsujPO() klasy ParserPylowChemii
    """
    FRAKCJA_RESP = 1
    FRAKCJA_WDYCH = 2
    CHWILOWKI = 3


# funkcje
# konwerter liczb rzymskich
def konwertujRzymskie(liczba):
    """
    Konwertuje liczby rzymskie do arabskich. Jeśli podany string nie był liczbą rzymską, zwraca go z powrotem
    """
    # konwersja najmniejszych liczb rzymskich jest wystarczająca dla zastosowań tego skryptu. Implementacja większych liczb zwiększyłaby prawdopodobieństwo niepożądanej konwersji skrótu literowego do liczby arabskiej
    rzymskie = {
        u'I': 1,
        u'V': 5,
        u'X': 10,
        u'L': 50,
        # u'C': 100,
        # u'D': 500,
        # u'M': 1000
    }

    krymskie = {
        u'IV': 4,
        u'IX': 9,
        u'XL': 40,
        # u'XC': 90,
        # u'CD': 400,
        # u'CM': 900
    }

    suma = 0
    dlugosc = len(liczba)
    licznik = 0
    for krymska in krymskie:
        if krymska in liczba:
            liczba = liczba.replace(krymska, u"")
            suma += krymskie[krymska]
            licznik += 2

    for rzymska in rzymskie:
        if rzymska in liczba:
            ile_razy = liczba.count(rzymska)
            suma += ile_razy * rzymskie[rzymska]
            licznik += ile_razy

    if licznik == dlugosc:
        return str(suma)
    else:
        return liczba


# globalny kontener na dane zczytane z plików tekstowych .vka, które powinny być czytane tylko raz i być ogólnodostępne
STALE_ZEWNETRZNE = {
    u"WYKONAWCA_POMIARU": None,
    u"METODY_POMIARÓW": None,
    u"SŁOWNIK_SKRÓTÓW": None,
    u"SŁOWNIK_ODMIAN": None
}


# TODO: dodać obsługę 'exists' jak w metodzie pobierzDRK()
def sparsujWykonawcePomiaru():
    """Parsuje z pliku tekstowego 'vkarter_stale.vka' informację o wykonawcy pomiaru"""
    wykonawca = []
    znacznik = u"*WYKONAWCA*"
    with codecs.open('vkarter_stale.vka', encoding='utf-8') as plik:
        for linia in plik:
            if znacznik in linia:
                linia = linia.rstrip(u'\n')
                wykonawca.append(linia.split(u':::')[1])

    assert len(wykonawca) > 0, "Nieudane parsowanie informacji o wykonawcy pomiaru. Brak pliku 'vkarter_stale.vka' lub plik nie zawiera danych w prawidlowym formacie"

    return u"".join(wykonawca)


# TODO: dodać obsługę 'exists' jak w metodzie pobierzDRK()
def sparsujMetodyPomiarow():
    """
    Parsuje z pliku tekstowego 'vkarter_stale.vka' informacje o metodach pomiarów i zapisuje jako krotkę MetodyPomiarow w kontenerze STALE_ZEWNETRZNE
    """
    metoda_halas = []
    metoda_pyl = []
    metoda_pyl_SiO2 = []
    metoda_chemia = []
    with codecs.open('vkarter_stale.vka', encoding='utf-8') as plik:
        for linia in plik:
            if u"*METODA_HAŁAS*" in linia:
                linia = linia.rstrip(u'\n')
                metoda_halas.append(linia.split(u':::')[1])
            elif u"*METODA_PYŁ*" in linia:
                linia = linia.rstrip(u'\n')
                metoda_pyl.append(linia.split(u':::')[1])
            elif u"*METODA_PYŁ_SiO2*" in linia:
                linia = linia.rstrip(u'\n')
                metoda_pyl_SiO2.append(linia.split(u':::')[1])
            elif u"*METODA_CHEMIA*" in linia:
                linia = linia.rstrip(u'\n')
                metoda_chemia.append(linia.split(u':::')[1])

    assert all(len(metoda) > 0 for metoda in [metoda_halas, metoda_pyl, metoda_pyl_SiO2, metoda_chemia]), u"Nieudane parsowanie informacji o metodzie pomiaru HALASU. Brak pliku 'vkarter_stale.vka' lub plik nie zawiera danych w prawidowym formacie"

    return MetodyPomiarow(u"".join(metoda_halas), u"".join(metoda_pyl), u"".join(metoda_pyl_SiO2), u"".join(metoda_chemia))


# TODO: dodać obsługę 'exists' jak w metodzie pobierzDRK()
def sparsujSkrotyDlaNazwyPliku():
    """
    Parsuje słownik skrótów dla GeneratoraNazwyPliku umieszczony w pliku tekstowym 'vkarter_skroty.vka'
    """
    slownik = []
    with codecs.open('vkarter_skroty.vka', encoding='utf-8') as plik:
        for linia in plik:
            linia = linia.rstrip(u'\n')
            slownik.append(linia.split(u':::'))

    slownik = dict(slownik)

    assert len(slownik) > 0, u"Nieudane parsowanie slownika skrotow dla nazw plikow. Brak pliku 'vkarter_skroty.vka' lub plik nie zawiera danych w prawidlowym formacie"

    return slownik


# TODO: dodać obsługę 'exists' jak w metodzie pobierzDRK()
def sparsujOdmianyNazwCzynnikow():
    """
    Parsuje słownik odmian nazw Czynników potrzebny do właściwego parsowania wartości p.o. umieszczony w pliku tekstowym 'vkarter_odmiany.vka'
    """
    slownik = []
    with codecs.open('vkarter_odmiany.vka', encoding='utf-8') as plik:
        for linia in plik:
            linia = linia.rstrip(u'\n')
            slownik.append(linia.split(u':::'))

    slownik = dict(slownik)

    assert len(slownik) > 0, u"Nieudane parsowanie slownika skrotow odmian nazw Czynników. Brak pliku 'vkarter_odmiany.vka' lub plik nie zawiera danych w prawidlowym formacie"

    return slownik


def sparsujPlikiVKA():
    """
    Funkcja wykonywana na początku skryptu, która parsuje wszystkie dane zgromadzone w zewnętrznych plikach tekstowych .vka i umieszcza je w globalnym słowniku
    """
    STALE_ZEWNETRZNE[u"WYKONAWCA_POMIARU"] = sparsujWykonawcePomiaru()
    STALE_ZEWNETRZNE[u"METODY_POMIARÓW"] = sparsujMetodyPomiarow()
    STALE_ZEWNETRZNE[u"SŁOWNIK_SKRÓTÓW"] = sparsujSkrotyDlaNazwyPliku()
    STALE_ZEWNETRZNE[u"SŁOWNIK_ODMIAN"] = sparsujOdmianyNazwCzynnikow()


# TYMCZASOWO - DO TESTÓW
def wyswietlKomunikat(komunikat, dl_kom=50):
    print
    print "*" * dl_kom
    print
    print komunikat.center(dl_kom)
    print
    print "*" * dl_kom
