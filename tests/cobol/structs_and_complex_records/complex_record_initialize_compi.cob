*> vybe-test: cobol/structs_and_complex_records/complex_record_initialize_compiles
*> origin: languages/cobol/tests/cobol/test_structs_and_complex_records.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-A PIC X(5) VALUE "AB".
   05 WS-B PIC 9(3) VALUE 9.
   05 WS-C PIC X(5) VALUE "CD".
PROCEDURE DIVISION.
    INITIALIZE WS-REC.
    STOP RUN.

