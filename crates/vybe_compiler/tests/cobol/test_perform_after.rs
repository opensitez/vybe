use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_perform_varying_after() {
    let output = run_prints(&p(
        r#"
01 WS-I PIC 9.
01 WS-J PIC 9.
"#,
        r#"
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 2
      AFTER WS-J FROM 1 BY 1 UNTIL WS-J > 3
        DISPLAY WS-I " " WS-J
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec![
        "1 1",
        "1 2",
        "1 3",
        "2 1",
        "2 2",
        "2 3"
    ]);
}

#[test]
fn test_perform_varying_after_3d() {
    compile_ok(&p(
        r#"
01 WS-I PIC 9.
01 WS-J PIC 9.
01 WS-K PIC 9.
"#,
        r#"
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 2
      AFTER WS-J FROM 1 BY 1 UNTIL WS-J > 2
      AFTER WS-K FROM 1 BY 1 UNTIL WS-K > 2
        DISPLAY WS-I WS-J WS-K
    END-PERFORM.
"#,
    ));
}
