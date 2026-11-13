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
    for life in (20, 35, 44.04, 44.08):
        inputs, profiles, flags, report = app.load_params_workbook(
            edited_workbook(life, active_profiles=True)
        )
        expected_months = app.operational_life_months(life)
        assert inputs["operational_life"] == float(life)
        assert report["expected_operating_months"] == expected_months
        assert all(len(profiles[key]) == expected_months for key in ("energy_price", "opex", "unavailability"))

    inputs, profiles, flags, report = app.load_params_workbook(
        edited_workbook(25, active_profiles=False)
    )
    assert all(not flags[key] for key in ("energy_price", "opex", "unavailability"))
    assert all(profiles[key] == [] for key in ("energy_price", "opex", "unavailability"))

    inputs, profiles, flags, report = app.load_params_workbook(
        edited_workbook(22.25, active_profiles=True, numeric_as_text=True)
    )
    assert inputs["operational_life"] == 22.25
    assert inputs["nominal_capex"] == 75.5

    verified = dict(inputs)
    verified["energy_price_profile"] = profiles["energy_price"] if flags["energy_price"] else None
    verified["opex_profile"] = profiles["opex"] if flags["opex"] else None
    verified["unavailability_profile"] = profiles["unavailability"] if flags["unavailability"] else None
    verified["delay_cumulative_curve"] = profiles["delay"] if flags["delay"] else None
    verified["capex_variation_cumulative_curve"] = profiles["capex_variation"] if flags["capex_variation"] else None
    round_trip = app.params_workbook(verified)
    reloaded = app.load_params_workbook(Upload(round_trip))[0]
    assert reloaded["operational_life"] == 22.25
    assert reloaded["nominal_capex"] == 75.5
    assert reloaded["systematic_eoh_error"] is False
    assert reloaded["systematic_eoh_error_std"] == 5.0

    assert app.operational_life_months(44.0) == 528
    assert app.operational_life_months(44.04) == 528
    assert app.operational_life_months(44.08) == 529
    generated = app.generate_flows(
        stochastic_capex=False,
        stochastic_delay=False,
        nominal_time_to_cod=24,
        operational_life=44.08,
        stochastic_eoh=False,
    )
    assert len(generated[4]) == 24 + 529
    assert all(len(series) == 24 + 529 for series in generated[4:-1])
    params = app.input_defaults()
    params["iterations"] = 5
    params["systematic_eoh_error"] = True
    params["systematic_eoh_error_std"] = 5.0
    with_error = app.run_monte_carlo(params)
    without_error = app.run_monte_carlo(dict(params, systematic_eoh_error=False))
    assert with_error.reference_eoh == without_error.reference_eoh
    assert with_error.reference_irr == without_error.reference_irr
    single_params = dict(params, stochastic_capex=False, stochastic_delay=False, stochastic_eoh=False)
    single = app.run_single(single_params)
    assert len(single) == 12 and single[-1] != 0.0
    single_book = load_workbook(io.BytesIO(app.single_iteration_workbook(single, single_params)), data_only=True)
    assert single_book["Iteration Results"]["A8"].value == "Systematic EOH Error"
    assert abs(single_book["Iteration Results"]["B8"].value - single[-1]) < 1e-12
    single_book.close()
    assert app.run_single(dict(single_params, systematic_eoh_error=False))[-1] == 0.0
    mc_book = load_workbook(io.BytesIO(app.monte_carlo_results_workbook(with_error, params)), data_only=True)
    curves = mc_book["Cumulative Curves"]
    expected = {"IRR (%)", "IRR Variation (bps)", "NPV (M€)", "PBP (y)", "EOH (h)", "EOH Variation (%)", "Time to COD (m)", "Delay (m)", "CAPEX (M€)", "CAPEX Variation (%)"}
    assert {curves.cell(1, c).value for c in range(1, curves.max_column + 1) if curves.cell(1, c).value} == expected
    mc_book.close()
    print("All Excel import regression checks passed.")


if __name__ == "__main__":
    run_checks()
