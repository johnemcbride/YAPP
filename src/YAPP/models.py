from io import BytesIO
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import re
from .status_level import StatusLevel


class ComputeStatus:

    status_date = None
    project_name = ""
    source_data_frame = pd.DataFrame()
    cleansed_data = pd.DataFrame()
    aggregates = {}

    date_columns = ["Baseline Start", "Baseline Finish",
                    "Actual Start", "Actual Finish", "Expected Finish"]

    def __init__(self, source_dataframe, status_date, project_name="Core Commercial Transformation"):
        self.source_data_frame = source_dataframe
        self.status_date = status_date
        self.project_name = project_name

        self.cleanse_input_data()
        self.add_level_0_summary_column(self.project_name)
        self.add_calculated_fields()
        self.compute_aggregates()
        print(self.aggregates["Level_5_Aggregation"]['actual_finish'])

    def add_level_0_summary_column(self, project_name):
        """
        Add a 'Level 0 Summary' column to cleansed_data, with the same value (project_name) in every row.
        Ensures it is inserted as the *first* column in your hierarchy.
        """
        if "Level 0 Summary" not in self.cleansed_data.columns:
            # Insert as the left-most column
            self.cleansed_data.insert(0, "Level 0 Summary", project_name)

    def aggregate_at_level(self, level, status_datetime, hierarchy_columns):

        df = self.cleansed_data

        group_obj = df.groupby(hierarchy_columns[:(level+1)])

        aggregation = group_obj.agg(
            baseline_start=("baseline_start", "min"),
            baseline_finish=("baseline_finish", "max"),
            actual_start=("actual_start", "min"),
            actual_finish=("actual_finish", "max"),
            actual_cost=("actual_cost", "sum"),
            expected_finish=("calculated_expected_finish", "max"),
            earned_value=("earned_value", "sum"),
            planned_value=("planned_value", "sum"),
            sum_baseline_duration=("baseline_duration", "sum"),
        ).reset_index()

        for col in ["baseline_start", "baseline_finish", "actual_start", "expected_finish", "actual_finish"]:
            aggregation[col] = pd.to_datetime(
                aggregation[col], errors="coerce")

        def compute_baseline_duration(row):
            if pd.notna(row["baseline_start"]) and pd.notna(row["baseline_finish"]):
                return np.busday_count(np.datetime64(row["baseline_start"].date(), 'D'),
                                       np.datetime64(row["baseline_finish"].date(), 'D')) + 1
            return 0

        def compute_expected_duration(row):
            if pd.notna(row["actual_start"]) and pd.notna(row["expected_finish"]):
                return np.busday_count(np.datetime64(row["actual_start"].date(), 'D'),
                                       np.datetime64(row["expected_finish"].date(), 'D')) + 1
            return 0

        aggregation["baseline_duration"] = aggregation.apply(
            compute_baseline_duration, axis=1)
        aggregation["expected_duration"] = aggregation.apply(
            compute_expected_duration, axis=1)

        aggregation["ev_over_pv"] = np.where(
            aggregation["planned_value"] > 0,
            aggregation["earned_value"] / aggregation["planned_value"],
            0
        )

        # actual percent complete

        aggregation["actual_percent_complete"] = np.where(
            aggregation["sum_baseline_duration"] > 0,
            aggregation["earned_value"] / aggregation["sum_baseline_duration"],
            0
        )

        aggregation["expected_percent_complete"] = np.where(
            aggregation["sum_baseline_duration"] > 0,
            aggregation["planned_value"] /
            aggregation["sum_baseline_duration"],
            0
        )

        aggregation["actual_finish"] = np.where(
            aggregation["actual_percent_complete"] == 1,
            aggregation["actual_finish"].dt.date,
            pd.NaT
        )

        def determine_status(row):
            if pd.isna(row["baseline_start"]) or pd.isna(row["baseline_finish"]):
                return "Unplanned"
            if pd.notna(row["actual_percent_complete"]) and row["actual_percent_complete"] == 1:
                return "Complete"
            if row["baseline_start"] <= status_datetime and pd.isna(row["actual_start"]):
                return "Late Starting"
            if row["baseline_start"] > status_datetime and pd.isna(row["actual_start"]):
                return "Not Due To Start"

            # Handle edge case: If planned_value is 0 but earned_value > 0, it should be "In Progress"
            if row["planned_value"] == 0 and row["earned_value"] > 0:
                return "In Progress"

            # Normal ev_over_pv logic
            if pd.notna(row["actual_start"]):  # Task has started
                if row["ev_over_pv"] >= 0.8:
                    return "In Progress"
                elif 0.3 <= row["ev_over_pv"] < 0.8:
                    return "Delayed"
                else:
                    return "Severely Delayed"

            return "Severely Delayed"  # Default fallback

        aggregation["status"] = aggregation.apply(determine_status, axis=1)
        return aggregation

    def cleanse_input_data(self):
        def date_columns_to_datetime(df, date_columns):
            df = df.copy()
            for col in date_columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
            return df

        self.cleansed_data = date_columns_to_datetime(
            self.source_data_frame, self.date_columns)

    def add_calculated_fields(self):

        self.cleansed_data["baseline_start"] = self.cleansed_data["Baseline Start"]
        self.cleansed_data["baseline_finish"] = self.cleansed_data["Baseline Finish"]
        self.cleansed_data["actual_start"] = self.cleansed_data["Actual Start"]
        self.cleansed_data["actual_finish"] = self.cleansed_data["Actual Finish"]
        self.cleansed_data["expected_finish"] = self.cleansed_data["Expected Finish"]

        del self.cleansed_data["Baseline Start"]
        del self.cleansed_data["Baseline Finish"]
        del self.cleansed_data["Actual Start"]
        del self.cleansed_data["Actual Finish"]
        del self.cleansed_data["Expected Finish"]

        def busday_count_wrapper(start, end, status_datetime):
            if pd.isna(start) or pd.isna(end):
                return 0
            start_np = np.datetime64(start.date(), 'D')
            end_np = np.datetime64(min(end, status_datetime).date(), 'D')
            return np.busday_count(start_np, end_np)

        # Baseline duration
        # Working days from baseline start to finish

        def row_baseline_duration(row):
            if pd.notna(row["baseline_start"]) and pd.notna(row["baseline_finish"]):
                return np.busday_count(
                    np.datetime64(row["baseline_start"].date(), "D"),
                    np.datetime64(row["baseline_finish"].date(), "D")
                ) + 1
            return 0

        self.cleansed_data["baseline_duration"] = self.cleansed_data.apply(
            row_baseline_duration, axis=1)

        # Calculated expected finish
        # Actual finish if set, else
        # If not actually finished,
        # either the expected finish which has been inputted or the max of baseline finsih, the actual finish or status date

        self.cleansed_data['calculated_expected_finish'] = self.cleansed_data.apply(
            lambda row: (
                row['actual_finish'] if pd.notna(row['actual_finish']) else
                row['expected_finish'] if pd.notna(row['expected_finish']) else
                max(row['baseline_finish'], row['actual_finish'],
                    pd.to_datetime(self.status_date))
            ),
            axis=1)

        # Planned value
        # Essentially the time elapsed since the baseline start, up to a cap of the finish or status date

        def compute_planned_value(row, config_date):
            """
            From Excel

            =IF(Config!$B$2<$G1352,0,IF($G1352 <>"",MAX(NETWORKDAYS($G1352,MIN(Config!$B$2,$H1352)),0),0))
            =IF(status_date < baseline_start, 0, if(baseline_start not null, Max(working days between(baseline start, min(status, baseline finish))),0) )

            Replicates:

                IF config_date < start_date:
                    0
                ELSE IF start_date not blank:
                    max(NETWORKDAYS(start_date, min(config_date, end_date)), 0)
                ELSE:
                    0
            """
            start_date = row["baseline_start"]
            end_date = row["baseline_finish"]

            # If start_date is NaT/None → 0
            if pd.isna(start_date):
                return 0

            # If config_date < start_date → 0
            if pd.to_datetime(config_date) < start_date:
                return 0

            busdays = busday_count_wrapper(
                start_date, end_date, pd.to_datetime(
                    config_date)
            )+1  # as it's the start of the start date to the end of the finish date
            # Clamp to 0 if negative
            return max(busdays, 0)

        self.cleansed_data["planned_value"] = self.cleansed_data.apply(
            lambda row: compute_planned_value(row, self.status_date),
            axis=1
        )

        # Actual cost
        # How long it has taken so far or took in entirety
        # (Probably could be simplified...?)

        def compute_actual_cost(row, config_date):
            """
            From Excel

            =IF(I1352<>"",MAX(NETWORKDAYS(I1352,MIN(R1352,Config!$B$2)),0),0)
            =IF(if(actual_start not null, Max(working days between(actual start, min(calculated expected finish, status date))),0) )


            """
            start_date = row["actual_start"]
            end_date = row["calculated_expected_finish"]

            # If start_date is NaT/None → 0
            if pd.isna(start_date):
                return 0

            busdays = busday_count_wrapper(
                start_date, end_date, pd.to_datetime(
                    config_date)
            )+1  # as it's the start of the start date to the end of the finish date
            # Clamp to 0 if negative
            return max(busdays, 0)

        self.cleansed_data["actual_cost"] = self.cleansed_data.apply(
            lambda row: compute_actual_cost(row, self.status_date),
            axis=1
        )

        # Expected total cost
        # Actual start to expected finsih
        def row_expected_total_cost(row):
            if pd.notna(row["actual_start"]) and pd.notna(row["calculated_expected_finish"]):
                return np.busday_count(
                    np.datetime64(row["actual_start"].date(), "D"),
                    np.datetime64(
                        row["calculated_expected_finish"].date(), "D")
                ) + 1
            return 0
        self.cleansed_data["expected_total_cost"] = self.cleansed_data.apply(
            row_expected_total_cost, axis=1)

        # Earned value
        # if baseline start set then actual cost/expected total cost * baseline duration

        self.cleansed_data['earned_value'] = np.where(
            pd.notna(self.cleansed_data["baseline_start"]),
            self.cleansed_data["actual_cost"] /
            self.cleansed_data["expected_total_cost"] *
            self.cleansed_data['baseline_duration'],
            0
        )

        calculated_date_columns = ["baseline_start", "baseline_finish",
                                   "actual_start", "actual_finish", "calculated_expected_finish"]
        for col in calculated_date_columns:
            self.cleansed_data[col] = pd.to_datetime(
                self.cleansed_data[col], errors="coerce")

    def get_hierarchy_column_names(self):
        return [
            col for col in self.cleansed_data.columns if re.match(r"Level \d+", col)]

    def compute_aggregates(self):
        max_level = len(self.get_hierarchy_column_names())
        for level_num in range(0, max_level+1):
            self.aggregates[f"Level_{level_num}_Aggregation"] = self.aggregate_at_level(
                level_num, pd.to_datetime(self.status_date), self.get_hierarchy_column_names())

    def get_root_status(self):
        return StatusLevel(self, 0)
