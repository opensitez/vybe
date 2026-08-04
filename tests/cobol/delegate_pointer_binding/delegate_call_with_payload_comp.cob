*> vybe-test: cobol/delegate_pointer_binding/delegate_call_with_payload_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE IS PROCEDURE-POINTER.
01 PAY PIC X(10) VALUE "DATA".
PROCEDURE DIVISION.
    CALL P USING PAY.
    STOP RUN.

