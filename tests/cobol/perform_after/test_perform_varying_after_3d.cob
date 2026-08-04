*> vybe-test: cobol/perform_after/test_perform_varying_after_3d
*> origin: languages/cobol/tests/cobol/test_perform_after.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9.
01 WS-J PIC 9.
01 WS-K PIC 9.
PROCEDURE DIVISION.

    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 2
      AFTER WS-J FROM 1 BY 1 UNTIL WS-J > 2
      AFTER WS-K FROM 1 BY 1 UNTIL WS-K > 2
        DISPLAY WS-I WS-J WS-K
    END-PERFORM.
    STOP RUN.

