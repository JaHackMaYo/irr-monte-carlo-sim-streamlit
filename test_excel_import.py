"""Regression checks for the complete Excel input import."""
import io
from openpyxl import load_workbook
import streamlit_app as app


class Upload:
    def __init__(self, data):
        self._data = data

    def getvalue(self):
        return self._data


def edited_workbook(life, active_profiles=True, numeric_as_text=False):
    data = app.params_workbook(app.input_defaults())
    workbook = load_workbook(io.BytesIO(data))
    inputs = workbook["Inputs"]
    for row in range(2, inputs.max_row + 1):
        key = inputs.cell(row, 1).value
        if key == "operational_life":
            inputs.cell(row, 2).value = str(life) if numeric_as_text else life
        elif key == "nominal_capex" and numeric_as_text:
            inputs.cell(row, 2).value = "75,5"
    profiles = workbook["Operating Profiles"]
    for column in range(1, 4):
        profiles.cell(2, column).value = active_profiles
    if not active_profiles:
        for row in range(5, profiles.max_row + 1):
            for column in range(2, 5):
                profiles.cell(row, column).value = None
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return Upload(output.getvalue())


def run_checks():
    for life in (20, 35):
        inputs, profiles, flags, report = app.load_params_workbook(
            edited_workbook(life, active_profiles=True)
        )
        assert inputs["operational_life"] == life
        assert report["expected_operating_months"] == life * 12
        assert all(len(profiles[key]) == life * 12 for key in ("energy_price", "opex", "unavailability"))

    inputs, profiles, flags, report = app.load_params_workbook(
        edited_workbook(25, active_profiles=False)
    )
    assert all(not flags[key] for key in ("energy_price", "opex", "unavailability"))
    assert all(profiles[key] == [] for key in ("energy_price", "opex", "unavailability"))

    inputs, profiles, flags, report = app.load_params_workbook(
        edited_workbook(22, active_profiles=True, numeric_as_text=True)
    )
    assert inputs["operational_life"] == 22
    assert inputs["nominal_capex"] == 75.5

    verified = dict(inputs)
    verified["energy_price_profile"] = profiles["energy_price"] if flags["energy_price"] else None
    verified["opex_profile"] = profiles["opex"] if flags["opex"] else None
    verified["unavailability_profile"] = profiles["unavailability"] if flags["unavailability"] else None
    verified["delay_cumulative_curve"] = profiles["delay"] if flags["delay"] else None
    verified["capex_variation_cumulative_curve"] = profiles["capex_variation"] if flags["capex_variation"] else None
    round_trip = app.params_workbook(verified)
    reloaded = app.load_params_workbook(Upload(round_trip))[0]
    assert reloaded["operational_life"] == 22
    assert reloaded["nominal_capex"] == 75.5
    print("All Excel import regression checks passed.")


if __name__ == "__main__":
    run_checks()
