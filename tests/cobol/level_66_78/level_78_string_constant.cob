*> vybe-test: cobol/level_66_78/level_78_string_constant
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 APP-NAME       VALUE "COBOL Application".
       78 APP-VERSION    VALUE "1.0.0".
       01 ws-title       PIC X(40).
       PROCEDURE DIVISION.
           MOVE APP-NAME TO ws-title
           DISPLAY ws-title
           STOP RUN.

