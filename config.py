# Definition of paths to files used

class path:
    # testing file with macro that only waits for x seconds
    # in practice the macro refreshes some reports, waits and runs again
    excel_with_macro ="testingwb.xlsm"
    macro_to_run ="'testingwb.xlsm'!Module1.test"

    # replace with
    # "\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\AdvancedEmailScheduler\\MainExcel_Development_v2.2.xlam"
    # "'\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\AdvancedEmailScheduler\\MainExcel_Development_v2.2.xlam'!Timer.EventMacro"
