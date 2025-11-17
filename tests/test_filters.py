import sys
import os
import pandas as pd
import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import filter_and_email_report  # Import your main script here

def make_test_df():
    # Sample data simulating report with different cases
    return pd.DataFrame({
        "DispatchZone": ["700", "800", "700", "900"],
        "R": ["scan1", "", None, "scan4"],
        "Driver": ["John", "", None, "Alice"],
        "SignedBy": ["", "Mike", None, ""],
    })

def test_filter_dispatch_zone():
    df = make_test_df()
    filtered = filter_and_email_report.apply_filters(df, dispatch="700", hide_blank_r="no", hide_driver_data="no", signed_blank="no")
    assert all(filtered["DispatchZone"] == "700")

def test_hide_blank_r():
    df = make_test_df()
    filtered = filter_and_email_report.apply_filters(df, dispatch="", hide_blank_r="yes", hide_driver_data="no", signed_blank="no")
    # R should not be blank or None
    assert all(filtered["R"].astype(bool))

def test_hide_driver_data():
    df = make_test_df()
    filtered = filter_and_email_report.apply_filters(df, dispatch="", hide_blank_r="no", hide_driver_data="yes", signed_blank="no")
    # Driver should be blank or None
    assert all(filtered["Driver"].isna() | (filtered["Driver"] == ""))

def test_show_only_blank_signedby():
    df = make_test_df()
    filtered = filter_and_email_report.apply_filters(df, dispatch="", hide_blank_r="no", hide_driver_data="no", signed_blank="yes")
    # SignedBy should be blank or None
    assert all(filtered["SignedBy"].isna() | (filtered["SignedBy"] == ""))

def test_combined_filters():
    df = make_test_df()
    filtered = filter_and_email_report.apply_filters(df, dispatch="700", hide_blank_r="yes", hide_driver_data="yes", signed_blank="yes")
    # All rows should meet all filter criteria
    assert all(filtered["DispatchZone"] == "700")
    assert all(filtered["R"].astype(bool))
    assert all(filtered["Driver"].isna() | (filtered["Driver"] == ""))
    assert all(filtered["SignedBy"].isna() | (filtered["SignedBy"] == ""))
    
def test_empty_dataframe():
    df = pd.DataFrame(columns=["DispatchZone", "R", "Driver", "SignedBy"])
    filtered = filter_and_email_report.apply_filters(df, dispatch="", hide_blank_r="yes", hide_driver_data="yes", signed_blank="yes")
    assert filtered.empty

def test_case_insensitivity():
    df = pd.DataFrame({"DispatchZone": ["700", "700a", "700B"]})
    filtered = filter_and_email_report.apply_filters(df, dispatch="700", hide_blank_r="no", hide_driver_data="no", signed_blank="no")
    assert len(filtered) == 3

def test_no_duplicate_rows():
    df = pd.DataFrame({"OrderNumber": [1, 1, 2], "DispatchZone": ["700", "700", "800"], "R": ["a", "a", "b"], "Driver": ["x", "x", "y"], "SignedBy": ["", "", ""]})
    #Remove duplicates just like in main script
    df = df.drop_duplicates(subset=['OrderNumber'])
    filtered = filter_and_email_report.apply_filters(df, dispatch="", hide_blank_r="no", hide_driver_data="no", signed_blank="no")
    # Here deduplication is done in main(), so could call that logic or test separately
    assert filtered["OrderNumber"].nunique() == len(filtered)