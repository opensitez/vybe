*> vybe-test: cobol/module_calls_program_linkage/call_by_reference_runtime_value_preserved
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5) VALUE 12345.
PROCEDURE DIVISION.
    CALL "M4" USING BY REFERENCE A
        ON EXCEPTION
            DISPLAY A
    END-CALL
    STOP RUN.

