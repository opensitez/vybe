*> vybe-test: cobol/program_id_options/program_id_recursive_basic
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. factorial IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-sub-result PIC 9(10) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(10).
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n <= 1
               MOVE 1 TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n
               CALL "factorial" USING lk-n ls-sub-result
               ADD 1 TO lk-n
               MULTIPLY lk-n BY ls-sub-result GIVING lk-result
           END-IF
           GOBACK.

