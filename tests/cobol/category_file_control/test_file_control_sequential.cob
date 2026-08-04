*> vybe-test: cobol/category_file_control/test_file_control_sequential
*> origin: languages/cobol/tests/cobol/test_category_file_control.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-SEQ.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SEQ-FILE ASSIGN TO "seq.dat"
           ORGANIZATION IS SEQUENTIAL
           ACCESS MODE IS SEQUENTIAL
           FILE STATUS IS WS-STAT.
       DATA DIVISION.
       FILE SECTION.
       FD SEQ-FILE.
       01 SEQ-REC PIC X(10).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-STAT PIC XX.
       PROCEDURE DIVISION.
           DISPLAY "SEQ FILE PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "SEQ FILE PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SEQ FILE PARSED"
        DISPLAY "FAIL: want [SEQ FILE PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

