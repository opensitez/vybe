use super::helpers::compile_ok;

#[test]
fn test_procedure_division_using() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-PARAM1 PIC 9(3).
01 LNK-PARAM2 PIC X(5).
PROCEDURE DIVISION USING LNK-PARAM1 LNK-PARAM2.
    DISPLAY LNK-PARAM1.
    DISPLAY LNK-PARAM2.
    GOBACK.
"#,
    );
}

#[test]
fn test_procedure_division_using_value() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-VAL PIC 9(3).
PROCEDURE DIVISION USING BY VALUE LNK-VAL.
    DISPLAY LNK-VAL.
    GOBACK.
"#,
    );
}

#[test]
fn test_procedure_division_returning() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-RET PIC 9(3).
PROCEDURE DIVISION RETURNING LNK-RET.
    MOVE 100 TO LNK-RET.
    GOBACK.
"#,
    );
}

#[test]
fn test_procedure_division_mixed_passing_modes() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-VALUE PIC 9(3).
01 LNK-TEXT PIC X(5).
PROCEDURE DIVISION USING BY VALUE LNK-VALUE BY REFERENCE LNK-TEXT.
    DISPLAY LNK-VALUE.
    DISPLAY LNK-TEXT.
    GOBACK.
"#,
    );
}

#[test]
fn test_procedure_division_with_returning_and_using() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-IN PIC 9(3) VALUE 0.
01 LNK-OUT PIC 9(3).
PROCEDURE DIVISION USING LNK-IN RETURNING LNK-OUT.
    COMPUTE LNK-OUT = LNK-IN + 1.
    GOBACK.
"#,
    );
}
