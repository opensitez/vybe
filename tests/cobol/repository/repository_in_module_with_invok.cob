*> vybe-test: cobol/repository/repository_in_module_with_invoke
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS StringUtil AS "StringUtil"
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-util  OBJECT REFERENCE StringUtil.
       01 ws-input PIC X(30) VALUE "hello world".
       01 ws-len   PIC 99.
       PROCEDURE DIVISION.
           COMPUTE ws-len = LENGTH(TRIM(ws-input))
           DISPLAY ws-len
           STOP RUN.

