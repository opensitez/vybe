*> vybe-test: cobol/program_id_options/mutual_recursion_even_odd
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. is-even IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       01 ls-sub-result PIC X VALUE "?".
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE "Y" TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
               CALL "is-odd" USING ls-n-minus-1 lk-result
           END-IF
           GOBACK.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. is-odd IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE "N" TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
               CALL "is-even" USING ls-n-minus-1 lk-result
           END-IF
           GOBACK.

