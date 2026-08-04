*> vybe-test: cobol/copy_replacing/copy_replacing_with_in_library_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY10.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC IN COMMON-LIB REPLACING OLD-NAME BY NEW-NAME.
    STOP RUN.

