*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_invoke_by_name_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 N PIC X(10) VALUE "M1".
PROCEDURE DIVISION.
    CALL "INVOKE-NAME" USING O N.
    STOP RUN.

