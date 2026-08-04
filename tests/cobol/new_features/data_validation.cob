*> vybe-test: cobol/new_features/data_validation
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. VALIDATE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT  PIC X(20) VALUE "12345".
01 WS-NAME   PIC X(20) VALUE "John Doe".
01 WS-AMOUNT PIC S9(5)V99 VALUE 150.50.
PROCEDURE DIVISION.
    IF WS-INPUT IS NUMERIC
        DISPLAY "Input is numeric"
    ELSE
        DISPLAY "Input is not numeric"
    END-IF.
    IF WS-NAME IS ALPHABETIC
        DISPLAY "Name is alpha"
    END-IF.
    IF WS-AMOUNT IS POSITIVE
        DISPLAY "Amount is positive"
    END-IF.
    IF WS-AMOUNT IS NOT ZERO
        DISPLAY "Amount is non-zero"
    END-IF.
    STOP RUN.

