*> vybe-test: cobol/standard_compliant_core/inspect_tallying_and_replacing_works_on_same_field
*> origin: languages/cobol/tests/cobol/test_standard_compliant_core.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(8) VALUE "ABABXABA".
01 WS-COUNT PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    INSPECT WS-TEXT TALLYING WS-COUNT FOR ALL "A".
    DISPLAY WS-COUNT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-COUNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "4"
        DISPLAY "FAIL: want [4] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

