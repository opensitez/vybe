*> vybe-test: cobol/source_format/source_format_free_with_compute_compiles
*> origin: languages/cobol/tests/cobol/test_source_format.rs
>>SOURCE FORMAT FREE
IDENTIFICATION DIVISION.
PROGRAM-ID. SF6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9.
PROCEDURE DIVISION.
    COMPUTE X = 1 + 1.
    STOP RUN.

