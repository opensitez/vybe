*> vybe-test: cobol/special_registers_detail/tally_after_inspect_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "AABABAB".
PROCEDURE DIVISION.
    INSPECT S TALLYING TALLY FOR ALL "A".
    STOP RUN.

