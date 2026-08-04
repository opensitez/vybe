*> vybe-test: cobol/program_id_options/program_id_recursive_countdown
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. countdown IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-next PIC 9(3) VALUE 0.
       LINKAGE SECTION.
       01 lk-n PIC 9(3).
       PROCEDURE DIVISION USING lk-n.
           DISPLAY lk-n
           IF lk-n > 0
               SUBTRACT 1 FROM lk-n GIVING ls-next
               CALL "countdown" USING ls-next
           END-IF
           GOBACK.

