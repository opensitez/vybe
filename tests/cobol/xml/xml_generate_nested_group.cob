*> vybe-test: cobol/xml/xml_generate_nested_group
*> origin: languages/cobol/tests/cobol/test_xml.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-person.
           05 ws-first-name  PIC X(15) VALUE "John".
           05 ws-last-name   PIC X(20) VALUE "Smith".
           05 ws-address.
               10 ws-street  PIC X(30) VALUE "123 Main St".
               10 ws-city    PIC X(20) VALUE "Springfield".
               10 ws-state   PIC XX    VALUE "IL".
               10 ws-zip     PIC X(10) VALUE "62701".
       01 ws-xml  PIC X(1000).
       01 ws-len  PIC 9(5).
       PROCEDURE DIVISION.
           XML GENERATE ws-xml FROM ws-person
               COUNT IN ws-len
           DISPLAY ws-xml(1:ws-len)
           STOP RUN.

