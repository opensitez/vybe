*> vybe-test: cobol/xml/xml_parse_extract_content
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-xml PIC X(100)
           VALUE "<employee><name>Bob</name><dept>IT</dept></employee>".
       01 ws-current-element PIC X(20).
       01 ws-name-val        PIC X(20).
       01 ws-dept-val        PIC X(20).
       PROCEDURE DIVISION.
           XML PARSE ws-xml
               PROCESSING PROCEDURE extract-data
           DISPLAY ws-name-val
           DISPLAY ws-dept-val
           STOP RUN.
       extract-data SECTION.
           EVALUATE XML-CODE
               WHEN "START-OF-ELEMENT"
                   MOVE XML-TEXT TO ws-current-element
               WHEN "CONTENT-CHARACTERS"
                   EVALUATE ws-current-element
                       WHEN "name" MOVE XML-TEXT TO ws-name-val
                       WHEN "dept" MOVE XML-TEXT TO ws-dept-val
                   END-EVALUATE
               WHEN OTHER CONTINUE
           END-EVALUATE.

