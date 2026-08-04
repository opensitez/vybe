*> vybe-test: cobol/scope_terminators/test_terminators_string_overflow
*> origin: languages/cobol/tests/cobol/test_scope_terminators.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(5) VALUE "HELLO".
01 WS-B PIC X(5) VALUE "WORLD".
01 WS-DST PIC X(5).
PROCEDURE DIVISION.

    STRING WS-A WS-B DELIMITED BY SIZE INTO WS-DST
        ON OVERFLOW
            DISPLAY "OVERFLOW"
    END-STRING.
    
    UNSTRING WS-DST DELIMITED BY SPACE INTO WS-A WS-B
        ON OVERFLOW
            DISPLAY "OVERFLOW"
    END-UNSTRING.
    STOP RUN.

