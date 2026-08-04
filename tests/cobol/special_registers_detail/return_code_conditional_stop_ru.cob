*> vybe-test: cobol/special_registers_detail/return_code_conditional_stop_run
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    IF RETURN-CODE > 0
        STOP RUN
    END-IF.
    DISPLAY "CONTINUED".
    STOP RUN.

