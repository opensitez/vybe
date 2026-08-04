*> vybe-test: cobol/paragraph_section_flow/section_contains_nested_perform
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM OUTER-SEC.
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "20"
        DISPLAY "FAIL: want [20] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
OUTER-SEC SECTION.
    PERFORM INNER-PARA.
    PERFORM INNER-PARA.
INNER-PARA.
    ADD 10 TO S.
    STOP RUN.

