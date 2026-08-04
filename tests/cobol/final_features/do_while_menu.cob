*> vybe-test: cobol/final_features/do_while_menu
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MENU.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CHOICE PIC 9(1) VALUE 0.
PROCEDURE DIVISION.
    PERFORM WITH TEST AFTER UNTIL WS-CHOICE = 9
        DISPLAY "1. Option A"
        DISPLAY "2. Option B"
        DISPLAY "9. Exit"
        MOVE 9 TO WS-CHOICE
    END-PERFORM.
    DISPLAY "Goodbye".
    STOP RUN.

