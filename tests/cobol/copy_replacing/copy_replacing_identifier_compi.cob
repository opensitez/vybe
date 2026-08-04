*> vybe-test: cobol/copy_replacing/copy_replacing_identifier_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY2.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.
    STOP RUN.

