*> vybe-test: cobol/special_registers_detail/tally_used_as_counter_then_displayed
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(15) VALUE "MISSISSIPPI    ".
PROCEDURE DIVISION.
    MOVE 0 TO TALLY.
    INSPECT S TALLYING TALLY FOR ALL "S".
    DISPLAY TALLY.
    STOP RUN.

