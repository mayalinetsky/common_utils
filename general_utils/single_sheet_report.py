"""
Interface for creating temporary, single-sheet reports that have several fields for each run (method+dataset)
"""
from typing import List, Dict

from ..errors import MethodNotFound
from ..core import utils
import itertools

from ..core.dataset.dataset_contents_query import DatasetContentsQuery
from ..core.dataset.dataset_file_client import DatasetFileClient
from ..core.run_query import RunQuery
from ..core.universal_imports import *
from ..database.controllers.database_client import DatabaseClient, Method
from ..database.models.accuracy_report import DATA_SET_NAME


# Main function
def create_report_from(methods: list, field_names: list, dtypes: list, func, sheet_name: str,
                       db_client: DatabaseClient):
    """
    Create and open a temporary excel report containing a table with an index "Data Set",
    and columns are multi-index with a first index from [field_names] and
    a second index from methods.
    On error, returns a string with the error. Else, returns as empty string.

    @param methods: list of method names
    @param field_names: list of names for the outputs of func
    @param dtypes: dtype of each field
    @param func: function that gets dataset instance, method and returns a tuple with the values of field_names
    @param sheet_name: string for sheet_name
    """
    try:
        report = get_report_df_from(methods, field_names, dtypes, func, db_client)
    except Exception as e:
        return str(e)

    utils.df_to_tmp_excel(report, sheet_name)
    return ""


def get_report_df_from(methods: List[str], field_names: list, dtypes: list, func, db_client: DatabaseClient):
    """
    Return a Dataframe with an index "Data Set",
    and columns are multi-index with a first index from [field_names] and
    a second index from methods.

    @param methods: list of method names
    @param field_names: list of names for the outputs of func
    @param dtypes: dtype of each field
    @param func: function that gets a RunQuery instance and returns a tuple with the values of field_names
    @param db_client: at the type hint suggests
    """
    all_methods = db_client.get_all_methods()

    errors = list(set(methods) - set(all_methods))
    if len(errors) > 0:
        raise MethodNotFound(', '.join(errors))

    all_datasets = db_client.get_all_datasets()

    multi_index_columns = pd.MultiIndex.from_product([field_names, methods], names=["Field", "Method"])
    dataset_index = pd.Index(data=all_datasets, name="Data Set")
    report = pd.DataFrame(data=[], index=dataset_index, columns=multi_index_columns)

    all_dataset_instances = [db_client.get_dataset(name) for name in all_datasets]
    all_file_clients = {d.name: DatasetFileClient(d.absolute_path) for d in all_dataset_instances}

    for method in methods:
        method_instance = db_client.get_method(method)
        field_values = _get_score_for_method_(method_instance, all_file_clients, func, field_names)

        report.loc[field_values[DATA_SET_NAME].values, list(
            itertools.product(field_names, [method]))] = field_values.iloc[:, 1:].values

    report.dropna(axis="index", how="all", inplace=True)  # drop datasets that did not run with any method

    col_to_dtype = {column: dtype for column, dtype in zip(field_names, dtypes)}
    report = report.astype(col_to_dtype)

    return report


def _get_score_for_method_(method: Method, dataset_name_to_file_client: Dict[str, DatasetFileClient],
                           score_func, score_names):
    """
    Return dataframe with the scores of [method]'s.
    Columns: DATA_SET_NAME, score_names
    Index: numeric

    Assumes method exist.
    @param method: existing method instance
    @param score_func: function that gets a RunQuery instance and returns a tuple with the values of score_names
    @param score_names: the name of the outputs of score_func
    """
    dataset_list = method.datasets

    data = []
    for d in dataset_list:
        file_client = dataset_name_to_file_client[d]
        run_query = RunQuery(DatasetContentsQuery(file_client), method.name)
        score = score_func(run_query)
        if isinstance(score, tuple):
            data.append([d] + list(score))
        else:
            data.append([d, score])

    analysis = pd.DataFrame(data=data,
                            columns=[DATA_SET_NAME] + score_names)
    return analysis

