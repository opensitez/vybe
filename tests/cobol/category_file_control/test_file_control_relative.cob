*> vybe-test: cobol/category_file_control/test_file_control_relative
*> origin: languages/cobol/tests/cobol/test_category_file_control.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-REL.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT REL-FILE ASSIGN TO "rel.dat"
           ORGANIZATION IS RELATIVE
           ACCESS MODE IS DYNAMIC
           RELATIVE KEY IS REL-KEY.
       DATA DIVISION.
       FILE SECTION.
       FD REL-FILE.
       01 REL-REC PIC X(10).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 REL-KEY PIC 9(4).
       PROCEDURE DIVISION.
           DISPLAY "REL FILE PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "REL FILE PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "REL FILE PARSED"
        DISPLAY "FAIL: want [REL FILE PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

