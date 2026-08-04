*> vybe-test: cobol/final_features/cond_88_range
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. RANGE88.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 9(3) VALUE 25.
   88 IS-CHILD  VALUE 0 THRU 12.
   88 IS-TEEN   VALUE 13 THRU 19.
   88 IS-ADULT  VALUE 20 THRU 120.
PROCEDURE DIVISION.
    IF IS-ADULT
        DISPLAY "Adult"
    END-IF.
    STOP RUN.

