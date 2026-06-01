import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Generator wyników zawodów")

st.title("🏆 Generator wyników zawodów")

POMIN = {
    "Tenis stołowy",
    "Piłka nożna",
    "Zabawy w wodzie",
    "Zestaw konkurencji"
}

CZASOWE = [
    "Bieg",
    "Pływanie",
    "Nordic walking",
    "Rower stacjonarny",
    "Wioślarstwo",
    "Wspinaczka",
    "Wyścig"
]

def czy_czasowa(nazwa):
    return any(sl in nazwa for sl in CZASOWE)

plik = st.file_uploader(
    "Wgraj plik Excel",
    type=["xlsx"]
)

if plik:

    st.success("Plik został wgrany")

    if st.button("Generuj wyniki"):

        xls = pd.ExcelFile(plik)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            for arkusz in xls.sheet_names:

                if arkusz in POMIN:
                    continue

                df = pd.read_excel(
                    plik,
                    sheet_name=arkusz
                )

                if "Wynik" not in df.columns:
                    continue

                df["Wynik"] = (
                    df["Wynik"]
                    .astype(str)
                    .str.replace(",", ".", regex=False)
                )

                df["Wynik"] = pd.to_numeric(
                    df["Wynik"],
                    errors="coerce"
                )

                df = df.dropna(subset=["Wynik"])

                if len(df) == 0:
                    continue

                czasowa = czy_czasowa(arkusz)

                wyniki_koncowe = []

                grupy = df.groupby([
                    "Płeć",
                    "Przedział wiekowy uczestnika",
                    "Stopień niepełnosprawności"
                ])

                for nazwa_grupy, grupa in grupy:

                    if czasowa:

                        grupa = grupa.sort_values(
                            "Wynik",
                            ascending=True
                        ).copy()

                        grupa["Miejsce"] = (
                            grupa["Wynik"]
                            .rank(
                                method="min",
                                ascending=True
                            )
                            .astype(int)
                        )

                        jednostka = "s"

                    else:

                        grupa = grupa.sort_values(
                            "Wynik",
                            ascending=False
                        ).copy()

                        grupa["Miejsce"] = (
                            grupa["Wynik"]
                            .rank(
                                method="min",
                                ascending=False
                            )
                            .astype(int)
                        )

                        jednostka = "m"

                    grupa = grupa[
                        grupa["Miejsce"] <= 3
                    ]

                    if len(grupa) == 0:
                        continue

                    plec, wiek, stopien = nazwa_grupy

                    grupa["Kategoria"] = (
                        f"{plec} | {wiek} | {stopien}"
                    )

                    grupa["Wynik"] = (
                        grupa["Wynik"]
                        .round(2)
                        .astype(str)
                        + f" {jednostka}"
                    )

                    wyniki_koncowe.append(grupa)

                if len(wyniki_koncowe) == 0:
                    continue

                wyniki = pd.concat(
                    wyniki_koncowe,
                    ignore_index=True
                )

                kolumny = [
                    "Miejsce",
                    "Kategoria",
                    "Wynik",
                    "id",
                    "Name",
                    "Płeć",
                    "Przedział wiekowy uczestnika",
                    "Stopień niepełnosprawności",
                    "Rodzaj niepełnosprawności",
                    "Sposób poruszania się uczestnika",
                    "Szkoła / organizacja",
                    "input_name",
                    "Email opiekuna/kontaktowy",
                    "Telefon"
                ]

                kolumny = [
                    k for k in kolumny
                    if k in wyniki.columns
                ]

                wyniki = wyniki[kolumny]

                wyniki.to_excel(
                    writer,
                    sheet_name=arkusz[:31],
                    index=False
                )

        output.seek(0)

        st.success("Wyniki wygenerowane")

        st.download_button(
            label="📥 Pobierz wyniki.xlsx",
            data=output,
            file_name="wyniki.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )