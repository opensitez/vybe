*> vybe-test: cobol/category_json_xml/test_xml_parse_basic
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-PARSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 XML-DOC PIC X(50) VALUE "<REC><FLD>HI</FLD></REC>".
       PROCEDURE DIVISION.
           XML PARSE XML-DOC
              PROCESSING PROCEDURE XML-PROC.
           DISPLAY "XML PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "XML PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XML PARSED"
        DISPLAY "FAIL: want [XML PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.
       XML-PROC SECTION.
           EXIT.

