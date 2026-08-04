*> vybe-test: cobol/new_features/batch_processing_pattern
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. BATCH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORDS PIC 9(5) VALUE 0.
01 WS-ERRORS  PIC 9(5) VALUE 0.
01 WS-SUCCESS PIC 9(5) VALUE 0.
01 WS-I       PIC 9(5) VALUE 0.
01 WS-MOD     PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        ADD 1 TO WS-RECORDS
        COMPUTE WS-MOD = FUNCTION MOD(WS-I 7)
        IF WS-MOD = 0
            ADD 1 TO WS-ERRORS
        ELSE
            ADD 1 TO WS-SUCCESS
        END-IF
    END-PERFORM.
    DISPLAY "Total Records: " WS-RECORDS.
    DISPLAY "Successful:    " WS-SUCCESS.
    DISPLAY "Errors:        " WS-ERRORS.
    STOP RUN.

