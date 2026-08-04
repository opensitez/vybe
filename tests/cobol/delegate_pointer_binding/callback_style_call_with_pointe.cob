*> vybe-test: cobol/delegate_pointer_binding/callback_style_call_with_pointer_args_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CALLBACK USAGE IS PROCEDURE-POINTER.
01 WS-ARG PIC X(10) VALUE "PAYLOAD".
PROCEDURE DIVISION.
    CALL "INVOKE-CALLBACK" USING WS-CALLBACK WS-ARG.
    STOP RUN.

