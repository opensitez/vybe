*> vybe-test: cobol/category_json_xml/test_xml_generate_basic
*> origin: languages/cobol/tests/cobol/test_category_json_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. XML-GEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 REC.
          05 FLD PIC X(5) VALUE "HELLO".
       01 XML-DOC PIC X(50).
       PROCEDURE DIVISION.
           XML GENERATE XML-DOC FROM REC.
           DISPLAY "XML GEN PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "XML GEN PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XML GEN PARSED"
        DISPLAY "FAIL: want [XML GEN PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

