*> vybe-test: cobol/data_group_level/data_group_filler_contributes_to_group_size
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 FILLER PIC X(2) VALUE "AB".
   05 DATA-PART PIC X(3) VALUE "CDE".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY REC.
    MOVE SPACES TO WS-VYBE-L
    STRING REC DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDE"
        DISPLAY "FAIL: want [ABCDE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

