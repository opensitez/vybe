*> vybe-test: cobol/value_all_forms/value_high_values_compared_compiles
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(4) VALUE HIGH-VALUES.
PROCEDURE DIVISION.
    IF S > "ZZZZ"
        DISPLAY "HIGH"
    ELSE
        DISPLAY "NOT HIGH"
    END-IF.
    STOP RUN.

