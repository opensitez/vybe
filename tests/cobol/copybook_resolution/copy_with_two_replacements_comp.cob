*> vybe-test: cobol/copybook_resolution/copy_with_two_replacements_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR10.
PROCEDURE DIVISION.
    COPY RECORD-DEF REPLACING OLD-A BY NEW-A OLD-B BY NEW-B.
    STOP RUN.

