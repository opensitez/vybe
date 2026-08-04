*> vybe-test: cobol/program_id_options/program_id_initial_and_recursive
*> origin: languages/cobol/tests/cobol/test_program_id_options.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. init-rec IS INITIAL RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-depth PIC 9 VALUE 0.
       LINKAGE SECTION.
       01 lk-max PIC 9.
       PROCEDURE DIVISION USING lk-max.
           IF ls-depth < lk-max
               ADD 1 TO ls-depth
               CALL "init-rec" USING lk-max
           ELSE
               DISPLAY ls-depth
           END-IF
           GOBACK.

