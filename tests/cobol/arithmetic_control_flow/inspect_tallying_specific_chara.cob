*> vybe-test: cobol/arithmetic_control_flow/inspect_tallying_specific_character_counts_total
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(8) VALUE "ABABXABA".
01 CNT PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT TXT TALLYING CNT FOR ALL "A".
    DISPLAY CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "04"
        DISPLAY "FAIL: want [04] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

