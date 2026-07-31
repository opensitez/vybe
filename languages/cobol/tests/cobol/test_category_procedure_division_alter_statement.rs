use super::helpers::{compile_ok, run_prints};

#[test]
fn alter_redirects_to_target_paragraph() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. AL1.
PROCEDURE DIVISION.
    ALTER ENTRY TO PROCEED TO TARGET.
    GO TO ENTRY.
ENTRY.
    DISPLAY "SOURCE".
    STOP RUN.
TARGET.
    DISPLAY "TARGET".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["TARGET"]);
}

#[test]
fn alter_reassigns_target_at_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. AL2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ws-flip PIC X VALUE "0".
PROCEDURE DIVISION.
    IF ws-flip = "0"
        ALTER ROUTER TO PROCEED TO ALPHA
    ELSE
        ALTER ROUTER TO PROCEED TO BETA
    END-IF
    GO TO ROUTER.
ROUTER.
    DISPLAY "ROUTER".
ALPHA.
    DISPLAY "ALPHA".
    STOP RUN.
BETA.
    DISPLAY "BETA".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["ALPHA"]);
}

#[test]
fn alter_multiple_targets_are_all_accepted() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
        PROGRAM-ID. AL3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ws-flag PIC X VALUE "A".
PROCEDURE DIVISION.
    ALTER ENTRY-1 TO PROCEED TO ROUTE-A.
    ALTER ENTRY-1 TO PROCEED TO ROUTE-B.
    GO TO ENTRY-1.
ENTRY-1.
    DISPLAY "ONE".
    GO TO ENDER.
ROUTE-A.
    DISPLAY "A-ROUTE".
ROUTE-B.
    DISPLAY "B-ROUTE".
    STOP RUN.
"ENdER" SECTION.
    DISPLAY "UNREACHED".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["ONE", "A-ROUTE"]);
}

#[test]
fn alter_chained_redirects_from_reachable_flow() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. AL4.
PROCEDURE DIVISION.
    ALTER ENTRY TO PROCEED TO MID.
    GO TO ENTRY.
ENTRY.
    DISPLAY "ENTRY".
    GO TO EXIT.
MID.
    DISPLAY "MID".
    ALTER ENTRY TO PROCEED TO END.
    GO TO ENTRY.
END.
    DISPLAY "END".
    STOP RUN.
EXIT.
    DISPLAY "EXIT".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["ENTRY", "MID", "END"]);
}

#[test]
fn alter_allows_hyphenated_paragraph_names() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. AL5.
       PROCEDURE DIVISION.
           ALTER START-POINT TO PROCEED TO ALT-PATH.
           GO TO START-POINT.
       START-POINT.
           DISPLAY "START".
           STOP RUN.
       ALT-PATH.
           DISPLAY "ALT-PATH".
           STOP RUN.
"#,
    );
}

#[test]
fn alter_parse_with_section_like_names() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. AL6.
       PROCEDURE DIVISION.
           ALTER ROUTE-ONE TO PROCEED TO ROUTE-TWO.
       ROUTE-ONE.
           GO TO END-POINT.
       ROUTE-TWO.
           DISPLAY "ROUTE-TWO".
       END-POINT.
           STOP RUN.
"#,
    );
}
