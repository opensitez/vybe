*> vybe-test: cobol/paragraph_section_flow/section_with_multiple_paragraphs
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM MY-SECT.
    STOP RUN.
MY-SECT SECTION.
PA.
    DISPLAY "PA".
    MOVE SPACES TO WS-VYBE-L
    STRING "PA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PA"
        DISPLAY "FAIL: want [PA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
PB.
    DISPLAY "PB".
    MOVE SPACES TO WS-VYBE-L
    STRING "PB" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "PB"
        DISPLAY "FAIL: want [PB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

