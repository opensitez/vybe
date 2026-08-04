*> vybe-test: cobol/delegate_pointer_binding/generic_pointer_and_delegate_slot_compile
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR USAGE IS POINTER.
01 WS-CALLBACK USAGE IS PROCEDURE-POINTER.
PROCEDURE DIVISION.
    SET WS-PTR TO NULL.
    DISPLAY "READY".
    STOP RUN.

