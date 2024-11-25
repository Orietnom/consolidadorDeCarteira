import pandas as pd
import os
import yfinance as yf
from datetime import datetime
import shutil

TRANSACTIONS_PATH = os.getcwd() + r"\Recursos\Planilhas B3\Movimentações"
NEGOTIATIONS_PATH = os.getcwd() + r"\Recursos\Planilhas B3\Negociações"

class GetSplit:

    def __init__(self) -> None:
       self._get_data()

    def _get_data(self):

        df_list = []
        for file in os.listdir(TRANSACTIONS_PATH):
            df_list.append(pd.read_excel(f"{TRANSACTIONS_PATH}\\{file}"))

        self.transactions_df = pd.concat(df_list, ignore_index=True).sort_values(by="Produto")

        for index, row in self.transactions_df.iterrows():
            if (row["Produto"][4:6] in ["12", "14", "15"]) and (row["Entrada/Saída"] == "Credito"):
                self.transactions_df.loc[index, 'Produto'] = row["Produto"][:4] + "11"
            else:
                ticker = row["Produto"][:6]
                self.transactions_df.loc[index, 'Produto'] = ticker.rstrip()
            if row["Produto"][:6] == "BBPO11":
                self.transactions_df.loc[index, 'Produto'] = "TVRI11"

    def verify_stock_split(self) -> pd.DataFrame:

        tickers = []
        stock_split = self.transactions_df[self.transactions_df["Movimentação"] == "Desdobro"]
        if not stock_split.empty:
            print("Houve desdobramento")
            for index, row in stock_split.iterrows():

                data = {
                    "ticker": row["Produto"],
                    "quantity": row["Quantidade"]
                }
                tickers.append(data)

            return tickers

        else:
            return []


class GetBonification:

    def __init__(self) -> None:
       self._get_data()

    def _get_data(self):

        df_list = []
        for file in os.listdir(TRANSACTIONS_PATH):
            df_list.append(pd.read_excel(f"{TRANSACTIONS_PATH}\\{file}"))

        self.transactions_df = pd.concat(df_list, ignore_index=True).sort_values(by="Produto")

        for index, row in self.transactions_df.iterrows():
            if (row["Produto"][4:6] in ["12", "14", "15"]) and (row["Entrada/Saída"] == "Credito"):
                self.transactions_df.loc[index, 'Produto'] = row["Produto"][:4] + "11"
            else:
                ticker = row["Produto"][:6]
                self.transactions_df.loc[index, 'Produto'] = ticker.rstrip()
            if row["Produto"][:6] == "BBPO11":
                self.transactions_df.loc[index, 'Produto'] = "TVRI11"

    def verify_bonification(self) -> pd.DataFrame:
        bonification_fraction_df = self.transactions_df[
            (self.transactions_df["Movimentação"] == "Bonificação em Ativos")
            | (self.transactions_df["Movimentação"] == "Fração em Ativos")
        ]
        if not bonification_fraction_df.empty:
            bonification_fraction_df.loc[bonification_fraction_df["Movimentação"] == "Fração em Ativos", 'Quantidade'] *= -1
            bonification_grouped_df = bonification_fraction_df.groupby(by="Produto").agg({"Quantidade": "sum"}).reset_index().copy()
            bonification_df = bonification_grouped_df[bonification_grouped_df["Quantidade"] > 0].copy()
            bonification_df.rename(columns={"Produto": "Código de Negociação"}, inplace=True)
            bonification_df.rename(columns={"Quantidade": "Quantidade Bonificação"}, inplace=True)
            return bonification_df


class Earns:
    def __init__(self) -> None:
       self._get_data()

    def _get_data(self):

        df_list = []
        for file in os.listdir(TRANSACTIONS_PATH):
            df_list.append(pd.read_excel(f"{TRANSACTIONS_PATH}\\{file}"))

        self.transactions_df = pd.concat(df_list, ignore_index=True).sort_values(by="Produto")

        for index, row in self.transactions_df.iterrows():
            if (row["Produto"][4:6] in ["12", "14", "15"]) and (row["Entrada/Saída"] == "Credito"):
                self.transactions_df.loc[index, 'Produto'] = row["Produto"][:4] + "11"
            else:
                ticker = row["Produto"][:6]
                self.transactions_df.loc[index, 'Produto'] = ticker.rstrip()
            if row["Produto"][:6] == "BBPO11":
                self.transactions_df.loc[index, 'Produto'] = "TVRI11"

    def get_earns(self) -> pd.DataFrame:
        earns_df = self.transactions_df[
            (self.transactions_df["Movimentação"] == "Dividendo")
            | (self.transactions_df["Movimentação"] == "Juros Sobre Capital Próprio")
            | (self.transactions_df["Movimentação"] == "Rendimento")
        ]
        grouped_earns = earns_df.groupby("Produto").agg({"Valor da Operação": "sum"})
        grouped_earns = grouped_earns.reset_index()
        return grouped_earns


class Negociacao:

    def __init__(self) -> None:
        df_fiis = pd.read_excel(os.getcwd() + "\\Recursos\\fundosListados.xlsx")
        self.fiis = df_fiis["Ticker"].tolist()


    def get_data(self):
        df_list = []
        for file in os.listdir(NEGOTIATIONS_PATH):
            df_list.append(pd.read_excel(f"{NEGOTIATIONS_PATH}\\{file}"))

        self.df = pd.concat(df_list, ignore_index=True).sort_values(by="Código de Negociação")

        for index, row in self.df.iterrows():
            if row["Código de Negociação"][-1] == "F":
                size = len(row["Código de Negociação"])
                self.df.loc[index, 'Código de Negociação'] = row["Código de Negociação"][0:(size-1)]
            if row["Código de Negociação"][:6] == "BBPO11":
                self.df.loc[index, 'Código de Negociação'] = "TVRI11"

        self.first_buy = self.df.groupby("Código de Negociação")["Data do Negócio"].min()
        print("OK")


    def create_wallet(self):

        buy = self.df[self.df["Tipo de Movimentação"] == "Compra"].copy()
        buy["Preço Total Compra"] = buy["Preço"] * buy["Quantidade"]
        grouped_buy = buy.groupby("Código de Negociação", as_index=False).agg(
            {
                'Quantidade': 'sum',
                'Preço Total Compra': 'sum'

            }
        )

        sell = self.df[self.df["Tipo de Movimentação"] == "Venda"].copy()
        sell["Preço Total Venda"] = sell["Preço"] * sell["Quantidade"]
        grouped_sell = sell.groupby("Código de Negociação", as_index=False).agg(
            {
                'Quantidade': 'sum',
                'Preço Total Venda': 'sum'

            }
        )

        grouped_buy.rename(columns={"Quantidade": "Quantidade Comprada"}, inplace=True)
        grouped_sell.rename(columns={"Quantidade": "Quantidade Vendida"}, inplace=True)

        wallet = pd.merge(grouped_buy, grouped_sell, on="Código de Negociação", how="left")
        wallet.fillna(0, inplace=True)

        wallet = wallet[(wallet["Quantidade Comprada"] - wallet["Quantidade Vendida"]) > 0]

        wallet["Tipo"] = wallet["Código de Negociação"].apply(lambda ticker: "FII" if ticker in self.fiis else "Ações")

        wallet = self.split_ticker(wallet)
        wallet = self.bonification(wallet)

        wallet["Quantidade"] = wallet["Quantidade Comprada"] - wallet["Quantidade Vendida"]
        wallet["Preço Médio"] = round(wallet["Preço Total Compra"] / wallet["Quantidade Comprada"], 2)

        fiis = wallet[wallet["Tipo"] == "FII"]
        acoes = wallet[wallet["Tipo"] == "Ações"]

        wallet, earns_df = self.income(wallet)
        wallet = self.variation(wallet)

        wallet["Posição"] = wallet.apply(self.calculate_position, axis=1)

        wallet = wallet[
            [
                "Código de Negociação",
                "Tipo",
                "Quantidade",
                "Preço Médio",
                "Variação",
                "Preço Total Compra",
                "Posição",
                "Preço de Fechamento",
                "Proventos Recebidos"
            ]
        ]
        wallet.rename(columns={"Código de Negociação": "Ticker"}, inplace=True)

        with pd.ExcelWriter('Carteira.xlsx', engine='openpyxl') as writer:
            wallet.to_excel(writer, sheet_name="Carteira", index=False)
            earns_df.to_excel(writer, sheet_name="Proventos", index=False)
        print(wallet)

    def calculate_position(self, row: pd.Series):
        if row["Variação"] > 1:
            variation = (row["Variação"]/100) + 1
            return round(row["Preço Total Compra"] * variation, 2)
        else:
            variation = (100 + row["Variação"])/100
            return round(row["Preço Total Compra"] * variation, 2)

    def split_ticker(self, wallet: pd.DataFrame) -> pd.DataFrame:
        ticker = GetSplit()
        tickers = ticker.verify_stock_split()
        if tickers:
            for ticker in tickers:
                for index, row in wallet.iterrows():
                    if row["Código de Negociação"] == ticker["ticker"]:
                        wallet.loc[index, "Quantidade Comprada"] = row["Quantidade Comprada"] + ticker["quantity"]

        return wallet

    def bonification(self, wallet: pd.DataFrame) -> pd.DataFrame:

        bonification = GetBonification()
        bonification_df = bonification.verify_bonification()
        merged_df = wallet.merge(bonification_df, on="Código de Negociação", how="left").fillna(0)
        merged_df["Quantidade Comprada Nova"] = merged_df["Quantidade Comprada"] + merged_df["Quantidade Bonificação"]
        merged_df = merged_df.drop(["Quantidade Comprada"], axis=1)
        merged_df.rename(columns={"Quantidade Comprada Nova": "Quantidade Comprada"}, inplace=True)
        return merged_df


    def income(self, wallet: pd.DataFrame) -> pd.DataFrame:
        earn = Earns()
        earns_df = earn.get_earns()
        earns_df["Valor da Operação"] = earns_df["Valor da Operação"].apply(lambda preco: round(preco, 2))
        earns_df.rename(columns={"Produto": "Código de Negociação", "Valor da Operação": "Proventos Recebidos"}, inplace=True)
        merged_wallet_earns = wallet.merge(right=earns_df, on="Código de Negociação", how="left").fillna(0)

        print(earns_df)
        return merged_wallet_earns, earns_df


    def variation(self, wallet: pd.DataFrame) -> pd.DataFrame:

        closed_price = []
        ticker_list = wallet["Código de Negociação"].tolist()

        for ticker in ticker_list:
            price = yf.Ticker(ticker + ".SA").info
            try:
                closed_price.append(price["previousClose"])
            except KeyError:
                closed_price.append(0.00)
                print("Preço de fechamento não encontrado")

        wallet["Preço de Fechamento"] = closed_price
        wallet["Variação"] = wallet.apply(self.calculate_variation, axis=1)

        return wallet

    def calculate_variation(self, row: pd.Series) -> float:
        variation = ((row["Preço de Fechamento"] / row["Preço Médio"]) - 1 ) * 100
        return round(variation, 2)


    def run(self):
        self.get_data()
        self.create_wallet()

