*> vybe-test: cobol/move_semantics_extended/move_high_values_to_alphanumeric_field_is_accepted
*> origin: languages/cobol/tests/cobol/test_move_semantics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(5) VALUE SPACES.
PROCEDURE DIVISION.

    MOVE HIGH-VALUES TO WS-TXT.
    STOP RUN.

