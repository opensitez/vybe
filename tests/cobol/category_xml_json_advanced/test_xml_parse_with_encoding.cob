*> vybe-test: cobol/category_xml_json_advanced/test_xml_parse_with_encoding
*> origin: languages/cobol/tests/cobol/test_category_xml_json_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-ENC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 XML-DOC PIC X(50) VALUE "<ROOT><A>1</A></ROOT>".
       PROCEDURE DIVISION.
           XML PARSE XML-DOC
              WITH ENCODING 1208
              PROCESSING PROCEDURE XML-PROC.
           DISPLAY "XML ENCODING PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "XML ENCODING PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XML ENCODING PARSED"
        DISPLAY "FAIL: want [XML ENCODING PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.
       XML-PROC SECTION.
           EXIT.

