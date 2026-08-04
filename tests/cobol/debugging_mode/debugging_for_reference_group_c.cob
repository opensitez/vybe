*> vybe-test: cobol/debugging_mode/debugging_for_reference_group_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG12.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL REFERENCES OF WS-A.
END DECLARATIVES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GROUP.
   05 WS-A PIC X.
   05 WS-B PIC X.
PROCEDURE DIVISION.
    STOP RUN.

