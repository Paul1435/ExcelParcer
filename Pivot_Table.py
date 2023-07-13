import pandas as pd


def create_pivot_table(source, index, values, aggfunc):
    return pd.pivot_table(source, index=index, aggfunc=aggfunc, values=values)

