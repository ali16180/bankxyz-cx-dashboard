"""Headless render check for dashboard.py.

Renders every page (plus the age-heatmap toggle and each Service Experience
touchpoint) and fails on any uncaught exception, so a metadata change can't
silently leave a page blank.

    uv run --with pandas,plotly,streamlit,openpyxl python smoke_test.py
"""

import sys

from streamlit.testing.v1 import AppTest

PAGES = ["Ringkasan", "Brand Image", "Branch Facilities",
         "Service Experience", "ATM Experience"]
TOUCHPOINTS = ["Customer Service", "Teller", "Security",
               "Customer Advisor", "Service Electronics"]
TIMEOUT = 180


def run(label, state=None):
    at = AppTest.from_file("dashboard.py", default_timeout=TIMEOUT)
    for k, v in (state or {}).items():
        at.session_state[k] = v
    at.run()
    if at.exception:
        print(f"FAIL  {label}")
        for e in at.exception:
            print(f"      {e.value}")
        return False
    charts = len(at.get("plotly_chart"))
    metrics = len(at.get("metric"))
    warns = len(at.warning)
    print(f"ok    {label:44s} charts={charts:2d} metrics={metrics:2d} "
          f"warnings={warns}")
    return True


def main():
    results = []
    for page in PAGES:
        results.append(run(page, {"nav_page": page}))

    for tp in TOUCHPOINTS:
        results.append(run(f"Service Experience -> {tp}",
                           {"nav_page": "Service Experience", "se_tp_keep": tp}))

    for page, key in (("Brand Image", "hm_bi"), ("Branch Facilities", "hm_bf"),
                      ("ATM Experience", "hm_atm")):
        results.append(run(f"{page} + heatmap usia",
                           {"nav_page": page, key: True}))

    # Filters narrow enough to exercise the small-sample paths.
    results.append(run("Ringkasan + filter provinsi",
                       {"nav_page": "Ringkasan", "f_prov": "DKI Jakarta"}))
    results.append(run("Branch Facilities + filter usia",
                       {"nav_page": "Branch Facilities",
                        "f_usia": "17 -19 tahun"}))

    print(f"\n{sum(results)}/{len(results)} render checks passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
