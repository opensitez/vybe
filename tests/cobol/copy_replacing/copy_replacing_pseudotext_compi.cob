*> vybe-test: cobol/copy_replacing/copy_replacing_pseudotext_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY3.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING ==OLD-NAME== BY ==NEW-NAME==.
    STOP RUN.

