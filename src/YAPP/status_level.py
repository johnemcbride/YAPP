from io import BytesIO
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import re
import json


class StatusLevel:
    def __init__(self, statusObj, level_num, index_num=0):
        self.level_num = level_num
        self.statusObj = statusObj
        self.index_num = index_num

    def get_description(self):
        return self.statusObj.aggregates["Level_"+str(self.level_num)+"_Aggregation"]['Level '+str(self.level_num)+' Summary'].iloc[self.index_num]

    def get_children(self):
        children = []
        children_df = self.statusObj.aggregates["Level_" +
                                                str(self.level_num+1)+"_Aggregation"]
        children_rows = children_df[children_df['Level ' +
                                                str(self.level_num)+' Summary'] == self.get_description()]
        indexes = children_rows.index
        for index in indexes:
            children.append(StatusLevel(
                self.statusObj, self.level_num + 1, index))
        return children


def make_getter(field_name):
    def getter(self):
        df = self.statusObj.aggregates[f"Level_{self.level_num}_Aggregation"]
        val = df[field_name].iloc[self.index_num]
        # convert if needed...
        if pd.notna(val) and hasattr(val, "date"):
            return val.date()
        return val
    return getter


def make_formatted_getter(field_name):
    def getter(self):
        df = self.statusObj.aggregates[f"Level_{self.level_num}_Aggregation"]
        val = df[field_name].iloc[self.index_num]
        # format if date
        if pd.notna(val) and hasattr(val, "strftime"):
            return val.strftime("%d/%m/%Y")
        if pd.isna(val):
            return ""
        # fallback

        return str(val)
    return getter


def make_percentage_getter(field_name):
    """
    Returns the numeric percentage value as e.g. 33.0 (float) if underlying 
    is 0.33, or None if NaN.
    """

    def getter(self):
        df = self.statusObj.aggregates[f"Level_{self.level_num}_Aggregation"]
        val = df[field_name].iloc[self.index_num]

        # If it's NaN (missing):
        if pd.isna(val):
            return None

        # Multiply by 100 to get e.g. 33.0 for 0.33
        return val * 100
    return getter


def make_percentage_getter_formatted(field_name):
    """
    Returns a string like '33.00%' for underlying 0.33, or '' if NaN.
    """

    def getter(self):
        df = self.statusObj.aggregates[f"Level_{self.level_num}_Aggregation"]
        val = df[field_name].iloc[self.index_num]

        if pd.isna(val):
            return ""

        # Format to e.g. 33.00%
        return f"{val * 100:.2f}%"
    return getter


# Suppose you know a list of columns up front:
columns = ["baseline_start", "baseline_finish",
           "actual_start", "actual_finish", "expected_finish"]

# Dynamically attach them to MyDynamicGetters:
for col in columns:
    # For example, we define "get_baseline_start" as a method:
    setattr(StatusLevel, f"get_{col}", make_getter(col))

    # And "get_baseline_start_formatted":
    setattr(StatusLevel, f"get_{col}_formatted", make_formatted_getter(col))

percent_columns = ["actual_percent_complete", "expected_percent_complete"]


# Dynamically attach them to MyDynamicGetters:
for col in percent_columns:
    # For example, we define "get_baseline_start" as a method:
    setattr(StatusLevel, f"get_{col}", make_percentage_getter(col))

    # And "get_baseline_start_formatted":
    setattr(StatusLevel, f"get_{col}_formatted",
            make_percentage_getter_formatted(col))


def make_string_getter(field_name):
    """
    Returns the numeric percentage value as e.g. 33.0 (float) if underlying 
    is 0.33, or None if NaN.
    """

    def getter(self):
        df = self.statusObj.aggregates[f"Level_{self.level_num}_Aggregation"]
        val = df[field_name].iloc[self.index_num]

        # If it's NaN (missing):
        if pd.isna(val):
            return ''

        return val
    return getter


string_columns = ["status"]


# Dynamically attach them to MyDynamicGetters:
for col in string_columns:
    # For example, we define "get_baseline_start" as a method:
    setattr(StatusLevel, f"get_{col}", make_string_getter(col))
