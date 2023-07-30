from typing import List, Callable

import pandas as pd


def get_index_centered_report_df_from(inputs: list, column_names: List[str], dtypes: list, func: Callable):
    """
    Return a Dataframe with an index from [datasets],
    and columns [column_names] of dtypes [dtypes], using results from [func].

    @param inputs: list of inputs.
    @param column_names: names for the outputs of func
    @param dtypes: dtype of each column
    @param func: function that gets an input and returns a tuple with the values of column_names
    """
    inputs_index = pd.Index(data=inputs)
    report = pd.DataFrame(data=[], index=inputs_index, columns=column_names)

    for dataset in datasets:
        field_values = func(dataset)

        report.loc[dataset, :] = field_values

    col_to_dtype = {column: dtype for column, dtype in zip(column_names, dtypes)}
    report = report.astype(col_to_dtype, errors="ignore")

    return report

