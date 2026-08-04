*> vybe-test: cobol/program_id_options/program_id_recursive_with_local_storage
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. rec-sum IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-partial PIC 9(8) VALUE 0.
       01 ls-n-dec   PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(8).
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE 0 TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-dec
               CALL "rec-sum" USING ls-n-dec ls-partial
               ADD lk-n ls-partial GIVING lk-result
           END-IF
           GOBACK.

