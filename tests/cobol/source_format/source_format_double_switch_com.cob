*> vybe-test: cobol/source_format/source_format_double_switch_compiles
*> origin: languages/cobol/tests/cobol/test_source_format.rs
>>SOURCE FORMAT FREE
IDENTIFICATION DIVISION.
PROGRAM-ID. SF8.
>>SOURCE FORMAT FIXED
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9.
>>SOURCE FORMAT FREE
PROCEDURE DIVISION.
    STOP RUN.

