from modules import collector, to_excel
import pandas as pd

def merge(df):
    
    df = df.drop_duplicates(subset="Placa")
    df = df[["Placa", "Vencimento"]]
    # df["Feito"] = pd.to_datetime(df["Feito"], dayfirst=True)
    # df["Vencimento"] = df["Feito"] + pd.Timedelta(days=15)
    # del df["Feito"]

    print(f"{df.to_string()}\n\n=================================\n")

    to_excel(df)


def main():
    data = collector()

    df = pd.DataFrame(
        data,
        columns=[
            "Placa",
            "Tipo",
            "Marca",
            "Modelo",
            "Tipo Frota",
            "Responsável",
            "Proprietario",
            "Feito",
            "Vencimento",
            "None",
        ],
    )

    df.to_csv("result.csv")

    # df = pd.read_csv("result.csv")

    merge(df)


if __name__ == "__main__":
    main()