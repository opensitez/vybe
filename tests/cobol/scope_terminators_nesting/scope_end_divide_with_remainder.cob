*> vybe-test: cobol/scope_terminators_nesting/scope_end_divide_with_remainder
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 17.
01 Q PIC 9(3) VALUE 0.
01 REM PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    DIVIDE 5 INTO A GIVING Q REMAINDER REM
    END-DIVIDE.
    STOP RUN.

