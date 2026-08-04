*> vybe-test: cobol/category_compute/test_compute_with_size_error_clause
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9. 01 B PIC 9 VALUE 0. 01 C PIC X VALUE SPACES. PROCEDURE DIVISION. COMPUTE A = 999 ON SIZE ERROR MOVE 'Y' TO C END-COMPUTE. DISPLAY C A. STOP RUN.

