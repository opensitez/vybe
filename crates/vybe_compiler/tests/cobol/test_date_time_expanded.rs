use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn accept_from_date_compiles() { compile_ok(&p("01 D PIC X(8).", "    ACCEPT D FROM DATE.")); }
#[test] fn accept_from_day_compiles() { compile_ok(&p("01 D PIC X(5).", "    ACCEPT D FROM DAY.")); }
#[test] fn accept_from_time_compiles() { compile_ok(&p("01 T PIC X(8).", "    ACCEPT T FROM TIME.")); }
#[test] fn accept_from_day_of_week_compiles() { compile_ok(&p("01 W PIC X(1).", "    ACCEPT W FROM DAY-OF-WEEK.")); }
#[test] fn function_current_date_move_compiles() { compile_ok(&p("01 CD PIC X(21).", "    MOVE FUNCTION CURRENT-DATE TO CD.")); }
#[test] fn display_current_date_compiles() { compile_ok(&p("", "    DISPLAY CURRENT-DATE.")); }
#[test] fn date_to_yyyymmdd_style_compiles() { compile_ok(&p("01 D PIC X(8).", "    ACCEPT D FROM DATE YYYYMMDD.")); }
#[test] fn date_to_yyddd_style_compiles() { compile_ok(&p("01 D PIC X(5).", "    ACCEPT D FROM DAY YYYYDDD.")); }
#[test] fn time_hms_style_compiles() { compile_ok(&p("01 T PIC X(8).", "    ACCEPT T FROM TIME.")); }
#[test] fn datetime_display_concat_compiles() { compile_ok(&p("01 D PIC X(8).\n01 T PIC X(8).", "    ACCEPT D FROM DATE.\n    ACCEPT T FROM TIME.\n    DISPLAY D T.")); }
#[test] fn datetime_compare_if_compiles() { compile_ok(&p("01 T PIC X(8).", "    ACCEPT T FROM TIME.\n    IF T > \"12000000\"\n        DISPLAY \"PM\"\n    END-IF.")); }
#[test] fn datetime_evaluate_branch_compiles() { compile_ok(&p("01 W PIC X(1).", "    ACCEPT W FROM DAY-OF-WEEK.\n    EVALUATE W\n        WHEN \"1\" DISPLAY \"MON\"\n        WHEN OTHER DISPLAY \"N\"\n    END-EVALUATE.")); }
#[test] fn date_store_and_move_compiles() { compile_ok(&p("01 D1 PIC X(8).\n01 D2 PIC X(8).", "    ACCEPT D1 FROM DATE.\n    MOVE D1 TO D2.")); }
#[test] fn time_store_and_move_compiles() { compile_ok(&p("01 T1 PIC X(8).\n01 T2 PIC X(8).", "    ACCEPT T1 FROM TIME.\n    MOVE T1 TO T2.")); }
#[test] fn current_date_to_group_compiles() { compile_ok(&p("01 TS.\n   05 Y PIC 9(4).\n   05 M PIC 9(2).\n   05 D PIC 9(2).", "    MOVE FUNCTION CURRENT-DATE(1:8) TO TS.")); }
#[test] fn date_format_pipeline_compiles() { compile_ok(&p("01 D PIC X(8).\n01 OUT PIC X(8).", "    ACCEPT D FROM DATE.\n    MOVE FUNCTION TRIM(D) TO OUT.")); }
#[test] fn date_intrinsic_length_compiles() { compile_ok(&p("01 D PIC X(8).\n01 L PIC 9(3).", "    ACCEPT D FROM DATE.\n    MOVE FUNCTION LENGTH(D) TO L.")); }
#[test] fn time_intrinsic_trim_compiles() { compile_ok(&p("01 T PIC X(8).\n01 O PIC X(8).", "    ACCEPT T FROM TIME.\n    MOVE FUNCTION TRIM(T) TO O.")); }
#[test] fn date_compare_branch_compiles() { compile_ok(&p("01 D PIC X(8).", "    ACCEPT D FROM DATE.\n    IF D > \"20250101\" DISPLAY \"NEW\" ELSE DISPLAY \"OLD\" END-IF.")); }
#[test] fn time_compare_branch_compiles() { compile_ok(&p("01 T PIC X(8).", "    ACCEPT T FROM TIME.\n    IF T < \"12000000\" DISPLAY \"AM\" ELSE DISPLAY \"PM\" END-IF.")); }
#[test] fn datetime_store_group_move_compiles() { compile_ok(&p("01 D PIC X(8).\n01 TS.\n   05 Y PIC 9(4).\n   05 M PIC 9(2).\n   05 DD PIC 9(2).", "    ACCEPT D FROM DATE.\n    MOVE D TO TS.")); }
#[test] fn current_date_slice_components_compiles() { compile_ok(&p("01 CD PIC X(21).\n01 Y PIC X(4).", "    MOVE FUNCTION CURRENT-DATE TO CD.\n    MOVE CD(1:4) TO Y.")); }
#[test] fn day_of_week_evaluate_full_branch_compiles() { compile_ok(&p("01 W PIC X(1).", "    ACCEPT W FROM DAY-OF-WEEK.\n    EVALUATE W\n        WHEN \"1\" DISPLAY \"MON\"\n        WHEN \"2\" DISPLAY \"TUE\"\n        WHEN OTHER DISPLAY \"X\"\n    END-EVALUATE.")); }
#[test] fn date_time_concat_display_compiles() { compile_ok(&p("01 D PIC X(8).\n01 T PIC X(8).\n01 DT PIC X(20).", "    ACCEPT D FROM DATE.\n    ACCEPT T FROM TIME.\n    STRING D DELIMITED BY SIZE T DELIMITED BY SIZE INTO DT.\n    DISPLAY DT.")); }
