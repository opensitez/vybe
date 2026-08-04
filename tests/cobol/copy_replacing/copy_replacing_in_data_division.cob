*> vybe-test: cobol/copy_replacing/copy_replacing_in_data_division_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY9.
DATA DIVISION.
WORKING-STORAGE SECTION.
    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.
PROCEDURE DIVISION.
    STOP RUN.

