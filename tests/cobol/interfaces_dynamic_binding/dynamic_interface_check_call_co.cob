*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_interface_check_call_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 R PIC 9.
PROCEDURE DIVISION.
    CALL "IMPLEMENTS" USING O R.
    STOP RUN.

