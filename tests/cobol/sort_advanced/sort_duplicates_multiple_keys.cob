*> vybe-test: cobol/sort_advanced/sort_duplicates_multiple_keys
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sort-wf ASSIGN TO "swf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sort-wf.
       01 sort-record.
           05 dept-code  PIC X(4).
           05 emp-name   PIC X(20).
           05 salary     PIC 9(7)V99.
       PROCEDURE DIVISION.
           SORT sort-wf
               ON ASCENDING KEY dept-code
               ON ASCENDING KEY emp-name
               WITH DUPLICATES IN ORDER
               USING "employees.dat"
               GIVING "sorted-employees.dat"
           STOP RUN.

