use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_terminators_if_evaluate_perform() {
    compile_ok(&p(
        r#"
01 WS-A PIC 9 VALUE 5.
01 WS-I PIC 9.
"#,
        r#"
    IF WS-A > 0
        EVALUATE WS-A
            WHEN 5
                PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
                    DISPLAY WS-I
                END-PERFORM
        END-EVALUATE
    END-IF.
"#,
    ));
}

#[test]
fn test_terminators_arithmetic_size_error() {
    compile_ok(&p(
        r#"
01 WS-A PIC 9 VALUE 9.
"#,
        r#"
    ADD 1 TO WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-ADD.
    
    SUBTRACT 1 FROM WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-SUBTRACT.
    
    MULTIPLY 2 BY WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-MULTIPLY.
    
    DIVIDE 2 INTO WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-DIVIDE.
    
    COMPUTE WS-A = WS-A * 2
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-COMPUTE.
"#,
    ));
}

#[test]
fn test_terminators_string_overflow() {
    compile_ok(&p(
        r#"
01 WS-A PIC X(5) VALUE "HELLO".
01 WS-B PIC X(5) VALUE "WORLD".
01 WS-DST PIC X(5).
"#,
        r#"
    STRING WS-A WS-B DELIMITED BY SIZE INTO WS-DST
        ON OVERFLOW
            DISPLAY "OVERFLOW"
    END-STRING.
    
    UNSTRING WS-DST DELIMITED BY SPACE INTO WS-A WS-B
        ON OVERFLOW
            DISPLAY "OVERFLOW"
    END-UNSTRING.
"#,
    ));
}

#[test]
fn test_terminator_period_closing() {
    compile_ok(&p(
        "01 WS-A PIC 9 VALUE 5.",
        r#"
    IF WS-A > 0
        DISPLAY "POS"
        IF WS-A = 5
            DISPLAY "FIVE".
"#,
    ));
}
