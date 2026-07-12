use super::helpers::compile_ok;

#[test]
fn test_local_storage_basics() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. LS-PROG.
DATA DIVISION.
LOCAL-STORAGE SECTION.
01 LS-NUM PIC 9(3) VALUE 100.
01 LS-STR PIC X(5) VALUE "HELLO".
01 LS-TABLE.
   05 LS-ITEM OCCURS 3 TIMES PIC 9(3).
PROCEDURE DIVISION.
    ADD 1 TO LS-NUM.
    DISPLAY LS-NUM.
    DISPLAY LS-STR.
    GOBACK.
"#,
    );
}

#[test]
fn test_local_storage_value_clause() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. LS-PROG.
DATA DIVISION.
LOCAL-STORAGE SECTION.
01 LS-X PIC 9(3) VALUE ZERO.
01 LS-Y PIC X(5) VALUE SPACES.
PROCEDURE DIVISION.
    DISPLAY LS-X.
    DISPLAY LS-Y.
    GOBACK.
"#,
    );
}
