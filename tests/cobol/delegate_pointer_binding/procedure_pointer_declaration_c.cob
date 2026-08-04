*> vybe-test: cobol/delegate_pointer_binding/procedure_pointer_declaration_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PPTR USAGE IS PROCEDURE-POINTER.
PROCEDURE DIVISION.
    DISPLAY "PPTR".
    STOP RUN.

