*> vybe-test: cobol/category_json_xml/test_json_generate_exception
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-GEN-EXC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 REC PIC X(5) VALUE "HELLO".
       01 JSON-DOC PIC X(2).
       PROCEDURE DIVISION.
           JSON GENERATE JSON-DOC FROM REC
              ON EXCEPTION DISPLAY "EXCEPTION CAUGHT"
              NOT ON EXCEPTION DISPLAY "SUCCESS".
           STOP RUN.

