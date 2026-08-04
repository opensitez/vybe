*> vybe-test: cobol/program_id_recursive/program_id_recursive_group_data
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 EMPLOYEE.
   05 EMP-ID PIC 9(6) VALUE 100001.
   05 EMP-NAME PIC X(20) VALUE "ALICE".
PROCEDURE DIVISION.
    DISPLAY EMP-ID.
    DISPLAY EMP-NAME.
    STOP RUN.

