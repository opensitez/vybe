*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_dispatch_returning_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 R PIC 9.
PROCEDURE DIVISION.
    INVOKE O M1 RETURNING R.
    STOP RUN.

