*> vybe-test: cobol/alter_stop/alter_sentence_target_is_honoured_compiles
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           ALTER entry-point TO PROCEED TO target-path
           GO TO entry-point
           STOP RUN.
       entry-point.
           GO TO default-path.
       default-path.
           DISPLAY "DEFAULT".
           STOP RUN.
       target-path.
           DISPLAY "TARGET".
           STOP RUN.

