from modules import collector, to_excel
import pandas as pd

def merge(df):
    
    df = df.drop_duplicates(subset="Placa")
    df = df[["Placa", "Feito"]]
    df["Feito"] = pd.to_datetime(df["Feito"], dayfirst=True)
    df["Vencimento"] = df["Feito"] + pd.Timedelta(days=15)
    del df["Feito"]

    print(f"{df.to_string()}\n\n")

    df.to_csv("result.csv")

    to_excel(df)


def main():
    data = collector()

    df = pd.DataFrame(
        data,
        columns=[
            "Placa",
            "Tipo",
            "Modelo",
            "Tipo Frota",
            "Responsável",
            "Feito",
            "None",
        ],
    )

    merge(df)


if __name__ == "__main__":
    main()