      *> ============================================================
      *> STUDENT RECORDS AND GRADE PROCESSING SYSTEM
      *> ============================================================
      *> University student information system: enrollment,
      *> grade recording, GPA calculation, transcript generation,
      *> academic standing, degree audit.
      *>
      *> Demonstrates: COBOL 2014 intrinsic functions (MEAN, MAX,
      *> TRIM, UPPER-CASE), PERFORM with TEST AFTER, complex
      *> nested tables, STRING/UNSTRING, INSPECT CONVERTING,
      *> report-style output with page breaks.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. STUDENT-RECORDS.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT STUDENT-FILE ASSIGN TO "students.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS STU-ID
               ALTERNATE RECORD KEY IS STU-SSN
                   WITH DUPLICATES
               FILE STATUS IS WS-STU-STATUS.

           SELECT ENROLLMENT-FILE ASSIGN TO "enrollments.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-ENR-STATUS.

           SELECT GRADE-FILE ASSIGN TO "grades.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-GRD-STATUS.

           SELECT TRANSCRIPT-FILE ASSIGN TO "transcripts.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT HONOR-ROLL ASSIGN TO "honor_roll.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT PROBATION-LIST ASSIGN TO "probation.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  STUDENT-FILE
           RECORD CONTAINS 600 CHARACTERS.
       01  STUDENT-RECORD.
           05  STU-ID              PIC X(10).
           05  STU-SSN             PIC X(11).
           05  STU-NAME.
               10  STU-LAST        PIC X(25).
               10  STU-FIRST       PIC X(20).
               10  STU-MIDDLE      PIC X(1).
           05  STU-DOB             PIC 9(8).
           05  STU-MAJOR           PIC X(6).
           05  STU-MINOR           PIC X(6).
           05  STU-LEVEL           PIC X(2).
               88  FRESHMAN        VALUE 'FR'.
               88  SOPHOMORE       VALUE 'SO'.
               88  JUNIOR          VALUE 'JR'.
               88  SENIOR          VALUE 'SR'.
               88  GRADUATE        VALUE 'GR'.
           05  STU-STATUS          PIC X(2).
               88  FULL-TIME       VALUE 'FT'.
               88  PART-TIME       VALUE 'PT'.
               88  WITHDRAWN       VALUE 'WD'.
               88  GRADUATED       VALUE 'GD'.
           05  STU-ADMIT-DATE      PIC 9(8).
           05  STU-GRAD-DATE       PIC 9(8).
           05  STU-CUMULATIVE-GPA  PIC 9(1)V9(4) COMP-3.
           05  STU-TOTAL-CREDITS   PIC 9(5)V99   COMP-3.
           05  STU-QUALITY-POINTS  PIC 9(7)V99   COMP-3.
           05  STU-FINANCIAL-HOLD  PIC X(1).
               88  FIN-HOLD        VALUE 'Y'.
               88  NO-FIN-HOLD     VALUE 'N'.
           05  STU-ACADEMIC-STAND  PIC X(2).
               88  GOOD-STANDING   VALUE 'GS'.
               88  PROBATION       VALUE 'PR'.
               88  SUSPENSION      VALUE 'SU'.
               88  HONOR-ROLL-STU  VALUE 'HR'.
               88  DEANS-LIST      VALUE 'DL'.
           05  STU-SEMESTER-TABLE.
               10  STU-SEM-ENTRY OCCURS 16 TIMES.
                   15  SEM-CODE    PIC X(6).
                   15  SEM-GPA     PIC 9(1)V9(4) COMP-3.
                   15  SEM-CREDITS PIC 9(3)V99   COMP-3.
                   15  SEM-ENROLLED PIC X(1).
           05  FILLER              PIC X(50).

       FD  ENROLLMENT-FILE
           RECORD CONTAINS 80 CHARACTERS.
       01  ENROLLMENT-RECORD.
           05  ENR-STUDENT-ID      PIC X(10).
           05  ENR-COURSE-CODE     PIC X(8).
           05  ENR-SECTION         PIC X(4).
           05  ENR-SEMESTER        PIC X(6).
           05  ENR-CREDITS         PIC 9(2)V99.
           05  ENR-STATUS          PIC X(2).
               88  ENR-ENROLLED    VALUE 'EN'.
               88  ENR-DROPPED     VALUE 'DR'.
               88  ENR-COMPLETED   VALUE 'CO'.
           05  FILLER              PIC X(46).

       FD  GRADE-FILE
           RECORD CONTAINS 80 CHARACTERS.
       01  GRADE-RECORD.
           05  GRD-STUDENT-ID      PIC X(10).
           05  GRD-COURSE-CODE     PIC X(8).
           05  GRD-SEMESTER        PIC X(6).
           05  GRD-LETTER-GRADE    PIC X(2).
           05  GRD-NUMERIC-GRADE   PIC 9(3)V99.
           05  GRD-CREDITS         PIC 9(2)V99.
           05  FILLER              PIC X(46).

       FD  TRANSCRIPT-FILE
           RECORD CONTAINS 132 CHARACTERS.
       01  TRANSCRIPT-LINE         PIC X(132).

       FD  HONOR-ROLL
           RECORD CONTAINS 80 CHARACTERS.
       01  HONOR-LINE              PIC X(80).

       FD  PROBATION-LIST
           RECORD CONTAINS 80 CHARACTERS.
       01  PROBATION-LINE          PIC X(80).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-STU-STATUS       PIC XX.
               88  STU-OK          VALUE '00'.
               88  STU-NOT-FOUND   VALUE '23'.
               88  STU-EOF         VALUE '10'.
           05  WS-ENR-STATUS       PIC XX.
               88  ENR-OK          VALUE '00'.
               88  ENR-EOF         VALUE '10'.
           05  WS-GRD-STATUS       PIC XX.
               88  GRD-OK          VALUE '00'.
               88  GRD-EOF         VALUE '10'.

       01  WS-GRADE-POINTS-TABLE.
           05  GP-ENTRY OCCURS 13 TIMES INDEXED BY GP-IDX.
               10  GP-LETTER       PIC X(2).
               10  GP-POINTS       PIC 9(1)V9(3).

       01  WS-WORK-FIELDS.
           05  WS-GRADE-POINTS     PIC 9(1)V9(3).
           05  WS-SEM-QUALITY-PTS  PIC 9(7)V99.
           05  WS-SEM-CREDITS      PIC 9(5)V99.
           05  WS-SEM-GPA          PIC 9(1)V9(4).
           05  WS-CURRENT-SEM      PIC X(6).
           05  WS-CURRENT-DATE     PIC 9(8).
           05  WS-SEM-IDX          PIC 9(2).
           05  WS-FULL-NAME        PIC X(47).
           05  WS-AGE              PIC 9(3).

       01  WS-COUNTERS.
           05  WS-STUDENTS-PROC    PIC 9(8) VALUE ZEROS.
           05  WS-HONOR-ROLL-CNT   PIC 9(6) VALUE ZEROS.
           05  WS-PROBATION-CNT    PIC 9(6) VALUE ZEROS.
           05  WS-GRADES-POSTED    PIC 9(8) VALUE ZEROS.
           05  WS-TOTAL-ENROLLED   PIC 9(8) VALUE ZEROS.

       01  WS-FORMATTED.
           05  WF-GPA              PIC 9.9999.
           05  WF-CREDITS          PIC ZZZ9.99.
           05  WF-GRADE            PIC ZZ9.99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-POST-GRADES
               UNTIL GRD-EOF
           PERFORM 3000-UPDATE-STANDINGS
           PERFORM 4000-GENERATE-TRANSCRIPTS
           PERFORM 5000-GENERATE-LISTS
           PERFORM 6000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           MOVE WS-CURRENT-DATE(1:6) TO WS-CURRENT-SEM
           OPEN I-O    STUDENT-FILE
           OPEN INPUT  ENROLLMENT-FILE
           OPEN INPUT  GRADE-FILE
           OPEN OUTPUT TRANSCRIPT-FILE
           OPEN OUTPUT HONOR-ROLL
           OPEN OUTPUT PROBATION-LIST
           PERFORM 1100-LOAD-GRADE-POINTS
           PERFORM 1200-READ-GRADE.

       1100-LOAD-GRADE-POINTS.
           MOVE 'A+' TO GP-LETTER(1)  MOVE 4.000 TO GP-POINTS(1)
           MOVE 'A ' TO GP-LETTER(2)  MOVE 4.000 TO GP-POINTS(2)
           MOVE 'A-' TO GP-LETTER(3)  MOVE 3.700 TO GP-POINTS(3)
           MOVE 'B+' TO GP-LETTER(4)  MOVE 3.300 TO GP-POINTS(4)
           MOVE 'B ' TO GP-LETTER(5)  MOVE 3.000 TO GP-POINTS(5)
           MOVE 'B-' TO GP-LETTER(6)  MOVE 2.700 TO GP-POINTS(6)
           MOVE 'C+' TO GP-LETTER(7)  MOVE 2.300 TO GP-POINTS(7)
           MOVE 'C ' TO GP-LETTER(8)  MOVE 2.000 TO GP-POINTS(8)
           MOVE 'C-' TO GP-LETTER(9)  MOVE 1.700 TO GP-POINTS(9)
           MOVE 'D+' TO GP-LETTER(10) MOVE 1.300 TO GP-POINTS(10)
           MOVE 'D ' TO GP-LETTER(11) MOVE 1.000 TO GP-POINTS(11)
           MOVE 'D-' TO GP-LETTER(12) MOVE 0.700 TO GP-POINTS(12)
           MOVE 'F ' TO GP-LETTER(13) MOVE 0.000 TO GP-POINTS(13).

       1200-READ-GRADE.
           READ GRADE-FILE
               AT END MOVE '10' TO WS-GRD-STATUS
           END-READ.

       2000-POST-GRADES.
           MOVE GRD-STUDENT-ID TO STU-ID
           READ STUDENT-FILE
               INVALID KEY
                   ADD 1 TO WS-GRADES-POSTED
                   PERFORM 1200-READ-GRADE
                   STOP RUN
           END-READ
           IF STU-OK
               PERFORM 2100-LOOKUP-GRADE-POINTS
               PERFORM 2200-UPDATE-STUDENT-GPA
               ADD 1 TO WS-GRADES-POSTED
               REWRITE STUDENT-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE
           END-IF
           PERFORM 1200-READ-GRADE.

       2100-LOOKUP-GRADE-POINTS.
           MOVE ZEROS TO WS-GRADE-POINTS
           PERFORM VARYING GP-IDX FROM 1 BY 1
               UNTIL GP-IDX > 13
               IF FUNCTION TRIM(GRD-LETTER-GRADE) =
                  FUNCTION TRIM(GP-LETTER(GP-IDX))
                   MOVE GP-POINTS(GP-IDX) TO WS-GRADE-POINTS
               END-IF
           END-PERFORM.

       2200-UPDATE-STUDENT-GPA.
           *> Add quality points for this course
           COMPUTE WS-SEM-QUALITY-PTS =
               GRD-CREDITS * WS-GRADE-POINTS
           ADD WS-SEM-QUALITY-PTS TO STU-QUALITY-POINTS
           ADD GRD-CREDITS TO STU-TOTAL-CREDITS

           *> Recalculate cumulative GPA
           IF STU-TOTAL-CREDITS > ZEROS
               COMPUTE STU-CUMULATIVE-GPA ROUNDED =
                   STU-QUALITY-POINTS / STU-TOTAL-CREDITS
           END-IF.

       3000-UPDATE-STANDINGS.
           MOVE LOW-VALUES TO STU-ID
           START STUDENT-FILE KEY >= STU-ID
               INVALID KEY STOP RUN
           END-START
           PERFORM 3100-STANDING-SCAN UNTIL STU-EOF.

       3100-STANDING-SCAN.
           READ STUDENT-FILE NEXT
               AT END MOVE '10' TO WS-STU-STATUS
           END-READ
           IF NOT STU-EOF
               ADD 1 TO WS-STUDENTS-PROC
               EVALUATE TRUE
                   WHEN STU-CUMULATIVE-GPA >= 3.800
                       MOVE 'DL' TO STU-ACADEMIC-STAND
                       ADD 1 TO WS-HONOR-ROLL-CNT
                   WHEN STU-CUMULATIVE-GPA >= 3.500
                       MOVE 'HR' TO STU-ACADEMIC-STAND
                       ADD 1 TO WS-HONOR-ROLL-CNT
                   WHEN STU-CUMULATIVE-GPA >= 2.000
                       MOVE 'GS' TO STU-ACADEMIC-STAND
                   WHEN STU-CUMULATIVE-GPA >= 1.500
                       MOVE 'PR' TO STU-ACADEMIC-STAND
                       ADD 1 TO WS-PROBATION-CNT
                   WHEN OTHER
                       MOVE 'SU' TO STU-ACADEMIC-STAND
                       ADD 1 TO WS-PROBATION-CNT
               END-EVALUATE

               *> Update class level based on credits
               EVALUATE TRUE
                   WHEN STU-TOTAL-CREDITS < 30
                       MOVE 'FR' TO STU-LEVEL
                   WHEN STU-TOTAL-CREDITS < 60
                       MOVE 'SO' TO STU-LEVEL
                   WHEN STU-TOTAL-CREDITS < 90
                       MOVE 'JR' TO STU-LEVEL
                   WHEN OTHER
                       MOVE 'SR' TO STU-LEVEL
               END-EVALUATE

               REWRITE STUDENT-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE
           END-IF.

       4000-GENERATE-TRANSCRIPTS.
           MOVE LOW-VALUES TO STU-ID
           START STUDENT-FILE KEY >= STU-ID
               INVALID KEY STOP RUN
           END-START
           PERFORM 4100-TRANSCRIPT-SCAN UNTIL STU-EOF.

       4100-TRANSCRIPT-SCAN.
           READ STUDENT-FILE NEXT
               AT END MOVE '10' TO WS-STU-STATUS
           END-READ
           IF NOT STU-EOF
               PERFORM 4200-PRINT-TRANSCRIPT
           END-IF.

       4200-PRINT-TRANSCRIPT.
           WRITE TRANSCRIPT-LINE FROM ALL '='
           MOVE SPACES TO TRANSCRIPT-LINE
           STRING 'OFFICIAL ACADEMIC TRANSCRIPT'
               DELIMITED SIZE INTO TRANSCRIPT-LINE
           WRITE TRANSCRIPT-LINE
           MOVE SPACES TO TRANSCRIPT-LINE
           STRING 'Student: '
                  FUNCTION TRIM(STU-FIRST) ' '
                  FUNCTION TRIM(STU-LAST)
                  '  ID: ' STU-ID
               DELIMITED SIZE INTO TRANSCRIPT-LINE
           WRITE TRANSCRIPT-LINE
           MOVE SPACES TO TRANSCRIPT-LINE
           STRING 'Major: ' STU-MAJOR
                  '  Level: ' STU-LEVEL
                  '  Status: ' STU-STATUS
               DELIMITED SIZE INTO TRANSCRIPT-LINE
           WRITE TRANSCRIPT-LINE
           MOVE STU-CUMULATIVE-GPA TO WF-GPA
           MOVE STU-TOTAL-CREDITS  TO WF-CREDITS
           MOVE SPACES TO TRANSCRIPT-LINE
           STRING 'Cumulative GPA: ' WF-GPA
                  '  Total Credits: ' WF-CREDITS
                  '  Standing: ' STU-ACADEMIC-STAND
               DELIMITED SIZE INTO TRANSCRIPT-LINE
           WRITE TRANSCRIPT-LINE
           WRITE TRANSCRIPT-LINE FROM ALL '-'.

       5000-GENERATE-LISTS.
           MOVE LOW-VALUES TO STU-ID
           START STUDENT-FILE KEY >= STU-ID
               INVALID KEY STOP RUN
           END-START
           WRITE HONOR-LINE FROM "HONOR ROLL / DEAN'S LIST"
           WRITE HONOR-LINE FROM ALL '-'
           WRITE PROBATION-LINE FROM "ACADEMIC PROBATION/SUSPENSION LIST"
           WRITE PROBATION-LINE FROM ALL '-'
           PERFORM 5100-LIST-SCAN UNTIL STU-EOF.

       5100-LIST-SCAN.
           READ STUDENT-FILE NEXT
               AT END MOVE '10' TO WS-STU-STATUS
           END-READ
           IF NOT STU-EOF
               MOVE STU-CUMULATIVE-GPA TO WF-GPA
               MOVE SPACES TO WS-FULL-NAME
               STRING FUNCTION TRIM(STU-FIRST) ' '
                      FUNCTION TRIM(STU-LAST)
                   DELIMITED SIZE INTO WS-FULL-NAME
               EVALUATE TRUE
                   WHEN HONOR-ROLL-STU OR DEANS-LIST
                       MOVE SPACES TO HONOR-LINE
                       STRING WS-FULL-NAME SPACES(4)
                              'GPA: ' WF-GPA
                              '  ' STU-ACADEMIC-STAND
                           DELIMITED SIZE INTO HONOR-LINE
                       WRITE HONOR-LINE
                   WHEN PROBATION OR SUSPENSION
                       MOVE SPACES TO PROBATION-LINE
                       STRING WS-FULL-NAME SPACES(4)
                              'GPA: ' WF-GPA
                              '  ' STU-ACADEMIC-STAND
                           DELIMITED SIZE INTO PROBATION-LINE
                       WRITE PROBATION-LINE
               END-EVALUATE
           END-IF.

       6000-PRINT-SUMMARY.
           DISPLAY "=== STUDENT RECORDS PROCESSING COMPLETE ==="
           DISPLAY "Students Processed : " WS-STUDENTS-PROC
           DISPLAY "Grades Posted      : " WS-GRADES-POSTED
           DISPLAY "Honor Roll         : " WS-HONOR-ROLL-CNT
           DISPLAY "On Probation       : " WS-PROBATION-CNT.

       9000-TERMINATE.
           CLOSE STUDENT-FILE ENROLLMENT-FILE GRADE-FILE
                 TRANSCRIPT-FILE HONOR-ROLL PROBATION-LIST.
