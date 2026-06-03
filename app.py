import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

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

        pdf_buffer = BytesIO()

        pdf = SimpleDocTemplate(pdf_buffer)

        styles = getSampleStyleSheet()

        pdf_elements = []

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

                for kategoria, grupa_pdf in wyniki.groupby("Kategoria"):

                    pdf_elements.append(
                        Paragraph(
                            f"{arkusz}",
                            styles["Heading2"]
                        )
                    )

                    pdf_elements.append(
                        Paragraph(
                            kategoria,
                            styles["Heading3"]
                        )
                    )

                    grupa_pdf = grupa_pdf.sort_values("Miejsce")

                    lines = []

                    for _, zawodnik in grupa_pdf.iterrows():

                        lines.append(
                            f'{int(zawodnik["Miejsce"])}. '
                            f'{zawodnik["Name"]} '
                            f'({zawodnik["id"]})'
                        )

                    tekst = "<br/>".join(lines)

                    pdf_elements.append(
                        Paragraph(
                            tekst,
                            styles["BodyText"]
                        )
                    )

                    pdf_elements.append(
                        Spacer(1, 12)
                    )

        pdf.build(pdf_elements)

        pdf_buffer.seek(0)

        output.seek(0)

        st.success("Wyniki wygenerowane")

        st.download_button(
            label="📥 Pobierz wyniki.xlsx",
            data=output,
            file_name="wyniki.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📄 Pobierz wolontariusze.pdf",
            data=pdf_buffer,
            file_name="wolontariusze.pdf",
            mime="application/pdf"
        )
