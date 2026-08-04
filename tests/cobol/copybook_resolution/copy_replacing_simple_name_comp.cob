*> vybe-test: cobol/copybook_resolution/copy_replacing_simple_name_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR5.
PROCEDURE DIVISION.
    COPY BASIC-BOOK REPLACING OLD-FIELD BY NEW-FIELD.
    STOP RUN.

