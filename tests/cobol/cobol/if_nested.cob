*> vybe-test: cobol/cobol/if_nested
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. IFNEST.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 9(3) VALUE 85.
PROCEDURE DIVISION.
    IF WS-SCORE >= 90
        DISPLAY "A"
    ELSE
        IF WS-SCORE >= 80
            DISPLAY "B"
        ELSE
            IF WS-SCORE >= 70
                DISPLAY "C"
            ELSE
                DISPLAY "F"
            END-IF
        END-IF
    END-IF.
    STOP RUN.

