*> vybe-test: cobol/program_id_options/program_id_recursive_fibonacci
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. fib IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-a PIC 9(10) VALUE 0.
       01 ls-b PIC 9(10) VALUE 0.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       01 ls-n-minus-2 PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(10).
       PROCEDURE DIVISION USING lk-n lk-result.
           EVALUATE TRUE
               WHEN lk-n = 0 MOVE 0 TO lk-result
               WHEN lk-n = 1 MOVE 1 TO lk-result
               WHEN OTHER
                   SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
                   SUBTRACT 2 FROM lk-n GIVING ls-n-minus-2
                   CALL "fib" USING ls-n-minus-1 ls-a
                   CALL "fib" USING ls-n-minus-2 ls-b
                   ADD ls-a ls-b GIVING lk-result
           END-EVALUATE
           GOBACK.

