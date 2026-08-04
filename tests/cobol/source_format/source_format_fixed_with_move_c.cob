*> vybe-test: cobol/source_format/source_format_fixed_with_move_compiles
*> origin: languages/cobol/tests/cobol/test_source_format.rs
>>SOURCE FORMAT FIXED
IDENTIFICATION DIVISION.
PROGRAM-ID. SF7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9.
PROCEDURE DIVISION.
    MOVE 1 TO X.
    STOP RUN.

