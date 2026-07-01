use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn perform_times_loop_compiles() {
    compile_ok(&p(
        "",
        "    PERFORM 3 TIMES\n        DISPLAY \"loop\"\n    END-PERFORM.",
    ));
}

#[test]
fn perform_until_loop_compiles() {
    compile_ok(&p(
        "01 WS-COUNT PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL WS-COUNT >= 3\n        ADD 1 TO WS-COUNT\n    END-PERFORM.",
    ));
}

#[test]
fn perform_varying_loop_compiles() {
    compile_ok(&p(
        "01 WS-COUNT PIC 9(2) VALUE 0.",
        "    PERFORM VARYING WS-COUNT FROM 1 BY 2 UNTIL WS-COUNT > 5\n        DISPLAY WS-COUNT\n    END-PERFORM.",
    ));
}

#[test]
fn perform_times_loop_runtime_prints_three_lines() {
    let output = run_prints(&p(
        "",
        "    PERFORM 3 TIMES\n        DISPLAY \"X\"\n    END-PERFORM.",
    ));
    assert_eq!(output, vec!["X", "X", "X"]);
}

#[test]
fn perform_until_loop_runtime_stops_on_boundary() {
    let output = run_prints(&p(
        "01 WS-COUNT PIC 9 VALUE 0.",
        "    PERFORM UNTIL WS-COUNT >= 3\n        ADD 1 TO WS-COUNT\n    END-PERFORM.\n    DISPLAY WS-COUNT.",
    ));
    assert_eq!(output, vec!["3"]);
}

#[test]
fn perform_varying_runtime_accumulates_expected_total() {
    let output = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 SUMV PIC 99 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 4\n        ADD I TO SUMV\n    END-PERFORM.\n    DISPLAY SUMV.",
    ));
    assert_eq!(output, vec!["10"]);
}

#[test]
fn nested_perform_loops_runtime_emits_expected_count() {
    let output = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 J PIC 9 VALUE 0.\n01 CNT PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2\n        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2\n            ADD 1 TO CNT\n        END-PERFORM\n    END-PERFORM.\n    DISPLAY CNT.",
    ));
    assert_eq!(output, vec!["4"]);
}
