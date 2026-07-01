use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn accept_date_from_system_compiles() { compile_ok(&p("01 WS-DATE PIC X(8).", "    ACCEPT WS-DATE FROM DATE.")); }
#[test] fn accept_day_from_system_compiles() { compile_ok(&p("01 WS-DAY PIC X(8).", "    ACCEPT WS-DAY FROM DAY.")); }
#[test] fn accept_time_from_system_compiles() { compile_ok(&p("01 WS-TIME PIC X(8).", "    ACCEPT WS-TIME FROM TIME.")); }
#[test] fn display_current_date_compiles() { compile_ok(&p("", "    DISPLAY CURRENT-DATE.")); }
#[test] fn display_current_time_compiles() { compile_ok(&p("", "    DISPLAY CURRENT-TIME.")); }
#[test] fn use_function_when_compiled() { compile_ok(&p("01 WS-DATE PIC X(8).", "    MOVE FUNCTION CURRENT-DATE TO WS-DATE.")); }
