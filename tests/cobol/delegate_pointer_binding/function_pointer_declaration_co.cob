*> vybe-test: cobol/delegate_pointer_binding/function_pointer_declaration_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FPTR USAGE IS FUNCTION-POINTER.
PROCEDURE DIVISION.
    DISPLAY "FPTR".
    STOP RUN.

