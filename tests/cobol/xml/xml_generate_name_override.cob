*> vybe-test: cobol/xml/xml_generate_name_override
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-cust-record.
           05 cust-id    PIC 9(5) VALUE 101.
           05 cust-name  PIC X(20) VALUE "John Doe".
           05 cust-email PIC X(30) VALUE "john@example.com".
       01 ws-xml PIC X(500).
       01 ws-len PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-cust-record
               COUNT IN ws-len
               NAMESPACE IS "http://example.com/cust"
               NAMESPACE-PREFIX IS "cust"
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

