*> vybe-test: cobol/delegate_pointer_binding/call_through_procedure_pointer_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE IS PROCEDURE-POINTER.
PROCEDURE DIVISION.
    CALL P.
    STOP RUN.

