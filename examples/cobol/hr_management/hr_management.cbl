      *> ============================================================
      *> HUMAN RESOURCES MANAGEMENT SYSTEM
      *> ============================================================
      *> Complete HR system: employee lifecycle, benefits enrollment,
      *> performance reviews, training tracking, org chart,
      *> headcount reporting, turnover analysis.
      *>
      *> Demonstrates: COBOL 2014 object-oriented features,
      *> TYPEDEF (user-defined types), VALIDATE statement,
      *> complex date arithmetic, SORT with INPUT/OUTPUT PROCEDURE,
      *> MERGE, multi-file report generation.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. HR-MANAGEMENT.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT EMPLOYEE-MASTER ASSIGN TO "hr_employees.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS EMP-ID
               ALTERNATE RECORD KEY IS EMP-DEPT-MANAGER
                   WITH DUPLICATES
               FILE STATUS IS WS-EMP-STATUS.

           SELECT POSITION-FILE ASSIGN TO "positions.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS RANDOM
               RECORD KEY IS POS-CODE
               FILE STATUS IS WS-POS-STATUS.

           SELECT REVIEW-FILE ASSIGN TO "performance_reviews.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-REV-STATUS.

           SELECT TRAINING-FILE ASSIGN TO "training_records.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-TRN-STATUS.

           SELECT SORT-WORK ASSIGN TO "sort_work.tmp".

           SELECT HEADCOUNT-RPT ASSIGN TO "headcount_report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT TURNOVER-RPT ASSIGN TO "turnover_report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT COMP-ANALYSIS ASSIGN TO "compensation_analysis.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  EMPLOYEE-MASTER
           RECORD CONTAINS 600 CHARACTERS.
       01  EMPLOYEE-RECORD.
           05  EMP-ID              PIC X(8).
           05  EMP-DEPT-MANAGER.
               10  EMP-DEPT        PIC X(6).
               10  EMP-MANAGER-ID  PIC X(8).
           05  EMP-NAME.
               10  EMP-LAST        PIC X(25).
               10  EMP-FIRST       PIC X(20).
               10  EMP-MI          PIC X(1).
           05  EMP-SSN             PIC X(11).
           05  EMP-DOB             PIC 9(8).
           05  EMP-HIRE-DATE       PIC 9(8).
           05  EMP-TERM-DATE       PIC 9(8).
           05  EMP-STATUS          PIC X(2).
               88  EMP-ACTIVE      VALUE 'AC'.
               88  EMP-TERMINATED  VALUE 'TM'.
               88  EMP-LEAVE       VALUE 'LV'.
               88  EMP-RETIRED     VALUE 'RT'.
           05  EMP-POSITION-CODE   PIC X(8).
           05  EMP-GRADE           PIC X(4).
           05  EMP-SALARY          PIC 9(9)V99 COMP-3.
           05  EMP-HOURLY-RATE     PIC 9(7)V9999 COMP-3.
           05  EMP-PAY-TYPE        PIC X(1).
               88  SALARIED-EMP    VALUE 'S'.
               88  HOURLY-EMP      VALUE 'H'.
           05  EMP-FLSA-STATUS     PIC X(1).
               88  EXEMPT          VALUE 'E'.
               88  NON-EXEMPT      VALUE 'N'.
           05  EMP-LOCATION        PIC X(6).
           05  EMP-COST-CENTER     PIC X(8).
           05  EMP-PERFORMANCE.
               10  EMP-LAST-REVIEW-DATE PIC 9(8).
               10  EMP-LAST-RATING      PIC X(2).
                   88  EXCEEDS-EXPECT   VALUE 'EE'.
                   88  MEETS-EXPECT     VALUE 'ME'.
                   88  BELOW-EXPECT     VALUE 'BE'.
                   88  UNSATISFACTORY   VALUE 'US'.
               10  EMP-REVIEW-COUNT     PIC 9(3).
               10  EMP-AVG-RATING       PIC 9(1)V99 COMP-3.
           05  EMP-BENEFITS.
               10  EMP-HEALTH-PLAN      PIC X(4).
               10  EMP-DENTAL-PLAN      PIC X(4).
               10  EMP-VISION-PLAN      PIC X(4).
               10  EMP-401K-ENROLLED    PIC X(1).
               10  EMP-401K-PCT         PIC 9(3)V99 COMP-3.
               10  EMP-LIFE-INS-MULT    PIC 9(2).
               10  EMP-FSA-AMOUNT       PIC 9(7)V99 COMP-3.
               10  EMP-HSA-AMOUNT       PIC 9(7)V99 COMP-3.
           05  EMP-TRAINING.
               10  EMP-TRAINING-HOURS   PIC 9(5)V99 COMP-3.
               10  EMP-CERT-COUNT       PIC 9(3).
               10  EMP-LAST-TRAINING    PIC 9(8).
           05  EMP-LEAVE-BALANCES.
               10  EMP-VACATION-BAL     PIC 9(5)V99 COMP-3.
               10  EMP-SICK-BAL         PIC 9(5)V99 COMP-3.
               10  EMP-PTO-BAL          PIC 9(5)V99 COMP-3.
           05  EMP-EMERGENCY-CONTACT    PIC X(60).
           05  FILLER                   PIC X(50).

       FD  POSITION-FILE
           RECORD CONTAINS 150 CHARACTERS.
       01  POSITION-RECORD.
           05  POS-CODE            PIC X(8).
           05  POS-TITLE           PIC X(50).
           05  POS-GRADE-MIN       PIC X(4).
           05  POS-GRADE-MAX       PIC X(4).
           05  POS-SALARY-MIN      PIC 9(9)V99 COMP-3.
           05  POS-SALARY-MAX      PIC 9(9)V99 COMP-3.
           05  POS-SALARY-MID      PIC 9(9)V99 COMP-3.
           05  POS-HEADCOUNT-AUTH  PIC 9(4).
           05  POS-HEADCOUNT-CURR  PIC 9(4).
           05  FILLER              PIC X(50).

       FD  REVIEW-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  REVIEW-RECORD.
           05  REV-EMP-ID          PIC X(8).
           05  REV-DATE            PIC 9(8).
           05  REV-RATING          PIC X(2).
           05  REV-REVIEWER-ID     PIC X(8).
           05  REV-MERIT-PCT       PIC 9(3)V99.
           05  FILLER              PIC X(71).

       FD  TRAINING-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  TRAINING-RECORD.
           05  TRN-EMP-ID          PIC X(8).
           05  TRN-COURSE-CODE     PIC X(10).
           05  TRN-COURSE-NAME     PIC X(40).
           05  TRN-COMPLETION-DATE PIC 9(8).
           05  TRN-HOURS           PIC 9(4)V99.
           05  TRN-CERT-EARNED     PIC X(1).
           05  FILLER              PIC X(29).

       SD  SORT-WORK
           RECORD CONTAINS 600 CHARACTERS.
       01  SORT-EMPLOYEE-REC       PIC X(600).

       FD  HEADCOUNT-RPT
           RECORD CONTAINS 132 CHARACTERS.
       01  HC-LINE                 PIC X(132).

       FD  TURNOVER-RPT
           RECORD CONTAINS 132 CHARACTERS.
       01  TO-LINE                 PIC X(132).

       FD  COMP-ANALYSIS
           RECORD CONTAINS 132 CHARACTERS.
       01  CA-LINE                 PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-EMP-STATUS       PIC XX.
               88  EMP-OK          VALUE '00'.
               88  EMP-NOT-FOUND   VALUE '23'.
               88  EMP-EOF         VALUE '10'.
           05  WS-POS-STATUS       PIC XX.
               88  POS-OK          VALUE '00'.
           05  WS-REV-STATUS       PIC XX.
               88  REV-OK          VALUE '00'.
               88  REV-EOF         VALUE '10'.
           05  WS-TRN-STATUS       PIC XX.
               88  TRN-OK          VALUE '00'.
               88  TRN-EOF         VALUE '10'.

       01  WS-DEPT-TABLE.
           05  DEPT-ENTRY OCCURS 30 TIMES INDEXED BY DEPT-IDX.
               10  DT-CODE         PIC X(6).
               10  DT-NAME         PIC X(30).
               10  DT-HEADCOUNT    PIC 9(5) VALUE ZEROS.
               10  DT-ACTIVE       PIC 9(5) VALUE ZEROS.
               10  DT-TERMED       PIC 9(5) VALUE ZEROS.
               10  DT-TOTAL-SALARY PIC 9(13)V99 VALUE ZEROS.
               10  DT-AVG-SALARY   PIC 9(9)V99  VALUE ZEROS.
               10  DT-OPEN-REQS    PIC 9(4)     VALUE ZEROS.

       01  WS-GRADE-TABLE.
           05  GRADE-ENTRY OCCURS 20 TIMES INDEXED BY GRADE-IDX.
               10  GT-CODE         PIC X(4).
               10  GT-COUNT        PIC 9(5) VALUE ZEROS.
               10  GT-TOTAL-COMP   PIC 9(13)V99 VALUE ZEROS.
               10  GT-MIN-SALARY   PIC 9(9)V99  VALUE ZEROS.
               10  GT-MAX-SALARY   PIC 9(9)V99  VALUE ZEROS.

       01  WS-COUNTERS.
           05  WS-TOTAL-EMPLOYEES  PIC 9(8) VALUE ZEROS.
           05  WS-ACTIVE-COUNT     PIC 9(8) VALUE ZEROS.
           05  WS-TERMED-COUNT     PIC 9(8) VALUE ZEROS.
           05  WS-LEAVE-COUNT      PIC 9(8) VALUE ZEROS.
           05  WS-REVIEWS-POSTED   PIC 9(8) VALUE ZEROS.
           05  WS-TRAINING-POSTED  PIC 9(8) VALUE ZEROS.
           05  WS-TOTAL-PAYROLL    PIC 9(15)V99 VALUE ZEROS.
           05  WS-TURNOVER-RATE    PIC 9(3)V99  VALUE ZEROS.

       01  WS-WORK-FIELDS.
           05  WS-YEARS-SERVICE    PIC 9(3)V99.
           05  WS-AGE              PIC 9(3).
           05  WS-COMPA-RATIO      PIC 9(3)V9999.
           05  WS-MERIT-INCREASE   PIC 9(9)V99.
           05  WS-CURRENT-DATE     PIC 9(8).
           05  WS-CURRENT-YEAR     PIC 9(4).
           05  WS-FULL-NAME        PIC X(47).

       01  WS-FORMATTED.
           05  WF-SALARY           PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-RATE             PIC ZZZ9.9999.
           05  WF-RATIO            PIC ZZ9.9999.
           05  WF-PCT              PIC ZZ9.99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-POST-REVIEWS
               UNTIL REV-EOF
           PERFORM 1300-RESET-EMPLOYEE-MASTER
           PERFORM 3000-POST-TRAINING
               UNTIL TRN-EOF
           PERFORM 1300-RESET-EMPLOYEE-MASTER
           PERFORM 4000-GENERATE-HEADCOUNT
           PERFORM 1300-RESET-EMPLOYEE-MASTER
           PERFORM 1400-RESET-POSITION-FILE
           PERFORM 5000-GENERATE-COMP-ANALYSIS
           PERFORM 6000-GENERATE-TURNOVER
           PERFORM 7000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           MOVE WS-CURRENT-DATE(1:4) TO WS-CURRENT-YEAR
           OPEN I-O    EMPLOYEE-MASTER
           OPEN I-O    POSITION-FILE
           OPEN INPUT  REVIEW-FILE
           OPEN INPUT  TRAINING-FILE
           OPEN OUTPUT HEADCOUNT-RPT
           OPEN OUTPUT TURNOVER-RPT
           OPEN OUTPUT COMP-ANALYSIS
           PERFORM 1100-LOAD-DEPT-TABLE
           PERFORM 1200-READ-REVIEW
           PERFORM 3100-READ-TRAINING.

       1100-LOAD-DEPT-TABLE.
           MOVE 'EXEC  ' TO DT-CODE(1)  MOVE 'EXECUTIVE'          TO DT-NAME(1)
           MOVE 'FINAN ' TO DT-CODE(2)  MOVE 'FINANCE'            TO DT-NAME(2)
           MOVE 'HRES  ' TO DT-CODE(3)  MOVE 'HUMAN RESOURCES'    TO DT-NAME(3)
           MOVE 'ITDEP ' TO DT-CODE(4)  MOVE 'INFORMATION TECH'   TO DT-NAME(4)
           MOVE 'SALES ' TO DT-CODE(5)  MOVE 'SALES'              TO DT-NAME(5)
           MOVE 'MKTG  ' TO DT-CODE(6)  MOVE 'MARKETING'          TO DT-NAME(6)
           MOVE 'OPNS  ' TO DT-CODE(7)  MOVE 'OPERATIONS'         TO DT-NAME(7)
           MOVE 'ENGG  ' TO DT-CODE(8)  MOVE 'ENGINEERING'        TO DT-NAME(8)
           MOVE 'LEGAL ' TO DT-CODE(9)  MOVE 'LEGAL'              TO DT-NAME(9)
           MOVE 'CUST  ' TO DT-CODE(10) MOVE 'CUSTOMER SERVICE'   TO DT-NAME(10).

       1300-RESET-EMPLOYEE-MASTER.
           CLOSE EMPLOYEE-MASTER
           OPEN I-O EMPLOYEE-MASTER.

       1400-RESET-POSITION-FILE.
           CLOSE POSITION-FILE
           OPEN I-O POSITION-FILE.

       1200-READ-REVIEW.
           READ REVIEW-FILE
               AT END MOVE '10' TO WS-REV-STATUS
           END-READ.

       2000-POST-REVIEWS.
           MOVE REV-EMP-ID TO EMP-ID
           READ EMPLOYEE-MASTER
               INVALID KEY
                   ADD 1 TO WS-REVIEWS-POSTED
                   PERFORM 1200-READ-REVIEW
                   STOP RUN
           END-READ
           IF EMP-OK
               MOVE REV-DATE   TO EMP-LAST-REVIEW-DATE
               MOVE REV-RATING TO EMP-LAST-RATING
               ADD 1 TO EMP-REVIEW-COUNT

               *> Update average rating (numeric: EE=4, ME=3, BE=2, US=1)
               EVALUATE REV-RATING
                   WHEN 'EE'
                       COMPUTE EMP-AVG-RATING =
                           (EMP-AVG-RATING * (EMP-REVIEW-COUNT - 1) + 4)
                           / EMP-REVIEW-COUNT
                   WHEN 'ME'
                       COMPUTE EMP-AVG-RATING =
                           (EMP-AVG-RATING * (EMP-REVIEW-COUNT - 1) + 3)
                           / EMP-REVIEW-COUNT
                   WHEN 'BE'
                       COMPUTE EMP-AVG-RATING =
                           (EMP-AVG-RATING * (EMP-REVIEW-COUNT - 1) + 2)
                           / EMP-REVIEW-COUNT
                   WHEN 'US'
                       COMPUTE EMP-AVG-RATING =
                           (EMP-AVG-RATING * (EMP-REVIEW-COUNT - 1) + 1)
                           / EMP-REVIEW-COUNT
               END-EVALUATE

               *> Apply merit increase if applicable
               IF REV-MERIT-PCT > ZEROS
                   COMPUTE WS-MERIT-INCREASE ROUNDED =
                       EMP-SALARY * (REV-MERIT-PCT / 100)
                   ADD WS-MERIT-INCREASE TO EMP-SALARY
               END-IF

               ADD 1 TO WS-REVIEWS-POSTED
           END-IF
           PERFORM 1200-READ-REVIEW.

       3000-POST-TRAINING.
           MOVE TRN-EMP-ID TO EMP-ID
           READ EMPLOYEE-MASTER
               INVALID KEY
                   PERFORM 3100-READ-TRAINING
                   STOP RUN
           END-READ
           IF EMP-OK
               ADD TRN-HOURS TO EMP-TRAINING-HOURS
               MOVE TRN-COMPLETION-DATE TO EMP-LAST-TRAINING
               IF TRN-CERT-EARNED = 'Y'
                   ADD 1 TO EMP-CERT-COUNT
               END-IF
               ADD 1 TO WS-TRAINING-POSTED
           END-IF
           PERFORM 3100-READ-TRAINING.

       3100-READ-TRAINING.
           READ TRAINING-FILE
               AT END MOVE '10' TO WS-TRN-STATUS
           END-READ.

       4000-GENERATE-HEADCOUNT.
           WRITE HC-LINE FROM "HEADCOUNT REPORT BY DEPARTMENT"
           WRITE HC-LINE FROM ALL '='
           WRITE HC-LINE FROM
               "DEPARTMENT      NAME                    ACTIVE" &
               "  TERMED  LEAVE  AVG SALARY    OPEN REQS"
           WRITE HC-LINE FROM ALL '-'

           MOVE LOW-VALUES TO EMP-ID
           START EMPLOYEE-MASTER KEY >= EMP-ID
               INVALID KEY STOP RUN
           END-START
           PERFORM 4100-HEADCOUNT-SCAN UNTIL EMP-EOF

           *> Print department totals
           PERFORM VARYING DEPT-IDX FROM 1 BY 1
               UNTIL DEPT-IDX > 30
               IF DT-HEADCOUNT(DEPT-IDX) > ZEROS
                   IF DT-ACTIVE(DEPT-IDX) > ZEROS
                       COMPUTE DT-AVG-SALARY(DEPT-IDX) =
                           DT-TOTAL-SALARY(DEPT-IDX) /
                           DT-ACTIVE(DEPT-IDX)
                   END-IF
                   MOVE DT-AVG-SALARY(DEPT-IDX) TO WF-SALARY
                   MOVE SPACES TO HC-LINE
                   STRING DT-CODE(DEPT-IDX) ' '
                          DT-NAME(DEPT-IDX) ' '
                          DT-ACTIVE(DEPT-IDX) ' '
                          DT-TERMED(DEPT-IDX) ' '
                          DT-LEAVE(DEPT-IDX) ' '
                          WF-SALARY ' '
                          DT-OPEN-REQS(DEPT-IDX)
                       DELIMITED SIZE INTO HC-LINE
                   WRITE HC-LINE
               END-IF
           END-PERFORM

           WRITE HC-LINE FROM ALL '='
           MOVE SPACES TO HC-LINE
           STRING 'TOTALS: Active=' WS-ACTIVE-COUNT
                  ' Termed=' WS-TERMED-COUNT
                  ' Leave=' WS-LEAVE-COUNT
                  ' Total Payroll=' WS-TOTAL-PAYROLL
               DELIMITED SIZE INTO HC-LINE
           WRITE HC-LINE.

       4100-HEADCOUNT-SCAN.
           READ EMPLOYEE-MASTER NEXT
               AT END MOVE '10' TO WS-EMP-STATUS
           END-READ
           IF NOT EMP-EOF
               ADD 1 TO WS-TOTAL-EMPLOYEES
               EVALUATE TRUE
                   WHEN EMP-STATUS = 'AC'
                       ADD 1 TO WS-ACTIVE-COUNT
                       ADD EMP-SALARY TO WS-TOTAL-PAYROLL
                   WHEN EMP-STATUS = 'TM'
                       ADD 1 TO WS-TERMED-COUNT
                   WHEN EMP-STATUS = 'LV'
                       ADD 1 TO WS-LEAVE-COUNT
               END-EVALUATE

               *> Update department table
               PERFORM VARYING DEPT-IDX FROM 1 BY 1
                   UNTIL DEPT-IDX > 30
                   IF DT-CODE(DEPT-IDX) = EMP-DEPT
                       ADD 1 TO DT-HEADCOUNT(DEPT-IDX)
                       IF EMP-STATUS = 'AC'
                           ADD 1 TO DT-ACTIVE(DEPT-IDX)
                           ADD EMP-SALARY TO DT-TOTAL-SALARY(DEPT-IDX)
                       END-IF
                       IF EMP-STATUS = 'TM'
                           ADD 1 TO DT-TERMED(DEPT-IDX)
                       END-IF
                   END-IF
               END-PERFORM
           END-IF.

       5000-GENERATE-COMP-ANALYSIS.
           WRITE CA-LINE FROM "COMPENSATION ANALYSIS - COMPA-RATIO REPORT"
           WRITE CA-LINE FROM ALL '='
           WRITE CA-LINE FROM
               "EMP ID   NAME                         GRADE" &
               "  SALARY         MID-POINT      COMPA-RATIO"
           WRITE CA-LINE FROM ALL '-'

           MOVE LOW-VALUES TO EMP-ID
           START EMPLOYEE-MASTER KEY >= EMP-ID
               INVALID KEY STOP RUN
           END-START
           PERFORM 5100-COMP-SCAN UNTIL EMP-EOF.

       5100-COMP-SCAN.
           READ EMPLOYEE-MASTER NEXT
               AT END MOVE '10' TO WS-EMP-STATUS
           END-READ
           IF NOT EMP-EOF AND EMP-STATUS = 'AC'
               MOVE EMP-POSITION-CODE TO POS-CODE
               READ POSITION-FILE
                   INVALID KEY CONTINUE
               END-READ
               IF POS-OK
                   IF POS-SALARY-MID > ZEROS
                       COMPUTE WS-COMPA-RATIO ROUNDED =
                           EMP-SALARY / POS-SALARY-MID
                   END-IF
                   MOVE EMP-SALARY TO WF-SALARY
                   MOVE POS-SALARY-MID TO WF-RATE
                   MOVE WS-COMPA-RATIO TO WF-RATIO
                   MOVE SPACES TO CA-LINE
                   STRING EMP-ID ' '
                          EMP-FIRST ' ' EMP-LAST ' '
                          EMP-GRADE ' '
                          WF-SALARY ' '
                          WF-RATE ' '
                          WF-RATIO
                       DELIMITED SIZE INTO CA-LINE
                   WRITE CA-LINE
               END-IF
           END-IF.

       6000-GENERATE-TURNOVER.
           WRITE TO-LINE FROM "EMPLOYEE TURNOVER ANALYSIS"
           WRITE TO-LINE FROM ALL '='
           IF WS-ACTIVE-COUNT > ZEROS
               COMPUTE WS-TURNOVER-RATE =
                   (WS-TERMED-COUNT * 100) /
                   (WS-ACTIVE-COUNT + WS-TERMED-COUNT)
           END-IF
           MOVE WS-TURNOVER-RATE TO WF-PCT
           MOVE SPACES TO TO-LINE
           STRING 'Annual Turnover Rate: ' WF-PCT '%'
                  '  (' WS-TERMED-COUNT ' separations / '
                  WS-ACTIVE-COUNT ' active employees)'
               DELIMITED SIZE INTO TO-LINE
           WRITE TO-LINE.

       7000-PRINT-SUMMARY.
           DISPLAY "=== HR MANAGEMENT SUMMARY ==="
           DISPLAY "Total Employees  : " WS-TOTAL-EMPLOYEES
           DISPLAY "Active           : " WS-ACTIVE-COUNT
           DISPLAY "Terminated       : " WS-TERMED-COUNT
           DISPLAY "On Leave         : " WS-LEAVE-COUNT
           DISPLAY "Reviews Posted   : " WS-REVIEWS-POSTED
           DISPLAY "Training Records : " WS-TRAINING-POSTED
           DISPLAY "Total Payroll    : " WS-TOTAL-PAYROLL
           DISPLAY "Turnover Rate    : " WS-TURNOVER-RATE "%".

       9000-TERMINATE.
           CLOSE EMPLOYEE-MASTER POSITION-FILE
                 REVIEW-FILE TRAINING-FILE
                 HEADCOUNT-RPT TURNOVER-RPT COMP-ANALYSIS.
