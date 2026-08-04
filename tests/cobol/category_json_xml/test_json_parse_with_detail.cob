*> vybe-test: cobol/category_json_xml/test_json_parse_with_detail
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. JSON-PARSE-DETAIL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 JSON-DOC PIC X(20) VALUE '{"A":"HI"}'.
       01 REC.
          05 A PIC X(2).
       PROCEDURE DIVISION.
           JSON PARSE JSON-DOC INTO REC
              WITH DETAIL
              ON EXCEPTION DISPLAY "EXC".
           DISPLAY A.
           STOP RUN.

