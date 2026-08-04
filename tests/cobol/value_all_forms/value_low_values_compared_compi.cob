*> vybe-test: cobol/value_all_forms/value_low_values_compared_compiles
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(4) VALUE LOW-VALUES.
PROCEDURE DIVISION.
    IF S < "AAAA"
        DISPLAY "LOW"
    ELSE
        DISPLAY "NOT LOW"
    END-IF.
    STOP RUN.

