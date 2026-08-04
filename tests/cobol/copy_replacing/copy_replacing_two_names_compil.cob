*> vybe-test: cobol/copy_replacing/copy_replacing_two_names_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY5.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING OLD-A BY NEW-A OLD-B BY NEW-B.
    STOP RUN.

