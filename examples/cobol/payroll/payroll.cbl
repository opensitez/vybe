      *> ============================================================
      *> PAYROLL PROCESSING SYSTEM
      *> ============================================================
      *> Processes employee payroll: gross pay, tax withholding,
      *> deductions, net pay. Produces pay stubs and summary report.
      *>
      *> Demonstrates: COMPUTE, EVALUATE, nested PERFORM,
      *> group items, REDEFINES, 88-level condition names,
      *> WORKING-STORAGE tables, STRING/UNSTRING.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PAYROLL-SYSTEM.
       AUTHOR. VYBE-EXAMPLES.
       DATE-WRITTEN. 2024-01-01.

       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SOURCE-COMPUTER. ANY-COMPUTER.
       OBJECT-COMPUTER. ANY-COMPUTER.

       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT EMPLOYEE-FILE ASSIGN TO "employees.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               ACCESS MODE IS SEQUENTIAL
               FILE STATUS IS WS-EMP-STATUS.
           SELECT PAYROLL-REPORT ASSIGN TO "payroll_report.txt"
               ORGANIZATION IS LINE SEQUENTIAL
               ACCESS MODE IS SEQUENTIAL
               FILE STATUS IS WS-RPT-STATUS.

       DATA DIVISION.
       FILE SECTION.

       FD  EMPLOYEE-FILE
           RECORD CONTAINS 120 CHARACTERS.
       01  EMPLOYEE-RECORD.
           05  EMP-ID              PIC 9(6).
           05  EMP-NAME.
               10  EMP-LAST-NAME   PIC X(20).
               10  EMP-FIRST-NAME  PIC X(15).
           05  EMP-DEPT            PIC X(4).
           05  EMP-PAY-TYPE        PIC X(1).
               88  SALARIED        VALUE 'S'.
               88  HOURLY          VALUE 'H'.
               88  CONTRACT        VALUE 'C'.
           05  EMP-PAY-RATE        PIC 9(7)V99.
           05  EMP-HOURS-WORKED    PIC 9(3)V99.
           05  EMP-ALLOWANCES      PIC 9(2).
           05  EMP-401K-PCT        PIC 9(2)V99.
           05  EMP-HEALTH-CODE     PIC X(1).
               88  HEALTH-SINGLE   VALUE 'S'.
               88  HEALTH-FAMILY   VALUE 'F'.
               88  HEALTH-NONE     VALUE 'N'.
           05  FILLER              PIC X(28).

       FD  PAYROLL-REPORT
           RECORD CONTAINS 132 CHARACTERS.
       01  REPORT-LINE             PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-FILE-STATUS.
           05  WS-EMP-STATUS       PIC XX VALUE SPACES.
               88  EMP-FILE-OK     VALUE '00'.
               88  EMP-FILE-EOF    VALUE '10'.
           05  WS-RPT-STATUS       PIC XX VALUE SPACES.

       01  WS-CONSTANTS.
           05  FEDERAL-TAX-RATE    PIC V9999 VALUE .2200.
           05  STATE-TAX-RATE      PIC V9999 VALUE .0550.
           05  FICA-RATE           PIC V9999 VALUE .0765.
           05  FICA-WAGE-BASE      PIC 9(6)  VALUE 160200.
           05  OVERTIME-RATE       PIC V9    VALUE 1.5.
           05  STANDARD-HOURS      PIC 9(3)  VALUE 40.
           05  HEALTH-SINGLE-AMT   PIC 9(4)V99 VALUE 0250.00.
           05  HEALTH-FAMILY-AMT   PIC 9(4)V99 VALUE 0750.00.
           05  ALLOWANCE-AMOUNT    PIC 9(4)V99 VALUE 0096.15.

       01  WS-CALCULATIONS.
           05  WS-REGULAR-PAY      PIC 9(8)V99 VALUE ZEROS.
           05  WS-OVERTIME-PAY     PIC 9(8)V99 VALUE ZEROS.
           05  WS-GROSS-PAY        PIC 9(8)V99 VALUE ZEROS.
           05  WS-FEDERAL-TAX      PIC 9(7)V99 VALUE ZEROS.
           05  WS-STATE-TAX        PIC 9(7)V99 VALUE ZEROS.
           05  WS-FICA-TAX         PIC 9(7)V99 VALUE ZEROS.
           05  WS-401K-DEDUCT      PIC 9(7)V99 VALUE ZEROS.
           05  WS-HEALTH-DEDUCT    PIC 9(7)V99 VALUE ZEROS.
           05  WS-TOTAL-DEDUCT     PIC 9(8)V99 VALUE ZEROS.
           05  WS-NET-PAY          PIC 9(8)V99 VALUE ZEROS.
           05  WS-TAXABLE-INCOME   PIC 9(8)V99 VALUE ZEROS.
           05  WS-OVERTIME-HOURS   PIC 9(3)V99 VALUE ZEROS.
           05  WS-REGULAR-HOURS    PIC 9(3)V99 VALUE ZEROS.

       01  WS-TOTALS.
           05  TOT-EMPLOYEES       PIC 9(5)    VALUE ZEROS.
           05  TOT-GROSS-PAY       PIC 9(10)V99 VALUE ZEROS.
           05  TOT-FEDERAL-TAX     PIC 9(9)V99 VALUE ZEROS.
           05  TOT-STATE-TAX       PIC 9(9)V99 VALUE ZEROS.
           05  TOT-FICA-TAX        PIC 9(9)V99 VALUE ZEROS.
           05  TOT-401K            PIC 9(9)V99 VALUE ZEROS.
           05  TOT-NET-PAY         PIC 9(10)V99 VALUE ZEROS.

       01  WS-DEPT-TABLE.
           05  DEPT-ENTRY OCCURS 10 TIMES
                         INDEXED BY DEPT-IDX.
               10  DEPT-CODE       PIC X(4).
               10  DEPT-NAME       PIC X(20).
               10  DEPT-GROSS      PIC 9(10)V99 VALUE ZEROS.
               10  DEPT-COUNT      PIC 9(4)     VALUE ZEROS.

       01  WS-REPORT-FIELDS.
           05  WS-CURRENT-DATE.
               10  WS-YEAR         PIC 9(4).
               10  WS-MONTH        PIC 9(2).
               10  WS-DAY          PIC 9(2).
           05  WS-PAGE-NUM         PIC 9(4) VALUE 1.
           05  WS-LINE-COUNT       PIC 9(3) VALUE 0.
           05  WS-LINES-PER-PAGE   PIC 9(3) VALUE 55.

       01  WS-FORMATTED-FIELDS.
           05  WF-GROSS-PAY        PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-NET-PAY          PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-FEDERAL-TAX      PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-STATE-TAX        PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-FICA-TAX         PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-401K             PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-HEALTH           PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-TOTAL-DEDUCT     PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-EMP-NAME         PIC X(36).

       01  WS-REPORT-LINES.
           05  RL-HEADER-1.
               10  FILLER          PIC X(40) VALUE SPACES.
               10  FILLER          PIC X(30)
                   VALUE "ACME CORPORATION PAYROLL REPORT".
               10  FILLER          PIC X(20) VALUE SPACES.
               10  FILLER          PIC X(6)  VALUE "PAGE: ".
               10  RL-PAGE-NUM     PIC ZZZ9.
               10  FILLER          PIC X(32) VALUE SPACES.
           05  RL-HEADER-2.
               10  FILLER          PIC X(40) VALUE SPACES.
               10  FILLER          PIC X(20)
                   VALUE "PAY PERIOD: BI-WEEKLY".
               10  FILLER          PIC X(72) VALUE SPACES.
           05  RL-COLUMN-HDR.
               10  FILLER          PIC X(6)  VALUE "EMP-ID".
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  FILLER          PIC X(36) VALUE "EMPLOYEE NAME".
               10  FILLER          PIC X(12) VALUE "GROSS PAY".
               10  FILLER          PIC X(12) VALUE "FED TAX".
               10  FILLER          PIC X(12) VALUE "STATE TAX".
               10  FILLER          PIC X(12) VALUE "FICA".
               10  FILLER          PIC X(12) VALUE "401K".
               10  FILLER          PIC X(12) VALUE "HEALTH".
               10  FILLER          PIC X(12) VALUE "NET PAY".
               10  FILLER          PIC X(4)  VALUE SPACES.
           05  RL-DETAIL.
               10  RL-EMP-ID       PIC 9(6).
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-EMP-NAME     PIC X(36).
               10  RL-GROSS        PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-FED-TAX      PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-STATE-TAX    PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-FICA         PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-401K         PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-HEALTH       PIC ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(2)  VALUE SPACES.
               10  RL-NET-PAY      PIC ZZZ,ZZZ,ZZ9.99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-EMPLOYEES
               UNTIL EMP-FILE-EOF
           PERFORM 3000-PRINT-TOTALS
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           OPEN INPUT  EMPLOYEE-FILE
           OPEN OUTPUT PAYROLL-REPORT
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           PERFORM 1100-LOAD-DEPT-TABLE
           PERFORM 1200-PRINT-HEADERS
           PERFORM 2100-READ-EMPLOYEE.

       1100-LOAD-DEPT-TABLE.
           MOVE 'ACCT' TO DEPT-CODE(1)
           MOVE 'ACCOUNTING'         TO DEPT-NAME(1)
           MOVE 'ENGG' TO DEPT-CODE(2)
           MOVE 'ENGINEERING'        TO DEPT-NAME(2)
           MOVE 'SALE' TO DEPT-CODE(3)
           MOVE 'SALES'              TO DEPT-NAME(3)
           MOVE 'MKTG' TO DEPT-CODE(4)
           MOVE 'MARKETING'          TO DEPT-NAME(4)
           MOVE 'HRES' TO DEPT-CODE(5)
           MOVE 'HUMAN RESOURCES'    TO DEPT-NAME(5)
           MOVE 'ITDP' TO DEPT-CODE(6)
           MOVE 'IT DEVELOPMENT'     TO DEPT-NAME(6)
           MOVE 'OPNS' TO DEPT-CODE(7)
           MOVE 'OPERATIONS'         TO DEPT-NAME(7)
           MOVE 'FNAN' TO DEPT-CODE(8)
           MOVE 'FINANCE'            TO DEPT-NAME(8)
           MOVE 'LGAL' TO DEPT-CODE(9)
           MOVE 'LEGAL'              TO DEPT-NAME(9)
           MOVE 'EXEC' TO DEPT-CODE(10)
           MOVE 'EXECUTIVE'          TO DEPT-NAME(10).

       1200-PRINT-HEADERS.
           MOVE WS-PAGE-NUM TO RL-PAGE-NUM
           WRITE REPORT-LINE FROM RL-HEADER-1
           WRITE REPORT-LINE FROM RL-HEADER-2
           MOVE ALL '-' TO REPORT-LINE
           WRITE REPORT-LINE
           WRITE REPORT-LINE FROM RL-COLUMN-HDR
           MOVE ALL '-' TO REPORT-LINE
           WRITE REPORT-LINE
           ADD 5 TO WS-LINE-COUNT.

       2000-PROCESS-EMPLOYEES.
           PERFORM 2200-CALCULATE-PAY
           PERFORM 2300-CALCULATE-TAXES
           PERFORM 2400-CALCULATE-DEDUCTIONS
           PERFORM 2500-CALCULATE-NET-PAY
           PERFORM 2600-UPDATE-TOTALS
           PERFORM 2700-PRINT-DETAIL
           PERFORM 2100-READ-EMPLOYEE.

       2100-READ-EMPLOYEE.
           READ EMPLOYEE-FILE INTO EMPLOYEE-RECORD
               AT END MOVE '10' TO WS-EMP-STATUS
           END-READ.

       2200-CALCULATE-PAY.
           EVALUATE TRUE
               WHEN SALARIED
                   COMPUTE WS-GROSS-PAY = EMP-PAY-RATE / 26
                   MOVE ZEROS TO WS-OVERTIME-PAY
                   MOVE ZEROS TO WS-REGULAR-PAY
               WHEN HOURLY
                   IF EMP-HOURS-WORKED > STANDARD-HOURS
                       COMPUTE WS-OVERTIME-HOURS =
                           EMP-HOURS-WORKED - STANDARD-HOURS
                       MOVE STANDARD-HOURS TO WS-REGULAR-HOURS
                   ELSE
                       MOVE EMP-HOURS-WORKED TO WS-REGULAR-HOURS
                       MOVE ZEROS TO WS-OVERTIME-HOURS
                   END-IF
                   COMPUTE WS-REGULAR-PAY =
                       WS-REGULAR-HOURS * EMP-PAY-RATE
                   COMPUTE WS-OVERTIME-PAY =
                       WS-OVERTIME-HOURS * EMP-PAY-RATE * OVERTIME-RATE
                   COMPUTE WS-GROSS-PAY =
                       WS-REGULAR-PAY + WS-OVERTIME-PAY
               WHEN CONTRACT
                   MOVE EMP-PAY-RATE TO WS-GROSS-PAY
                   MOVE ZEROS TO WS-OVERTIME-PAY
                   MOVE ZEROS TO WS-REGULAR-PAY
               WHEN OTHER
                   MOVE ZEROS TO WS-GROSS-PAY
           END-EVALUATE.

       2300-CALCULATE-TAXES.
           *> Taxable income after allowances
           COMPUTE WS-TAXABLE-INCOME =
               WS-GROSS-PAY - (EMP-ALLOWANCES * ALLOWANCE-AMOUNT)
           IF WS-TAXABLE-INCOME < ZEROS
               MOVE ZEROS TO WS-TAXABLE-INCOME
           END-IF

           COMPUTE WS-FEDERAL-TAX ROUNDED =
               WS-TAXABLE-INCOME * FEDERAL-TAX-RATE
           COMPUTE WS-STATE-TAX ROUNDED =
               WS-TAXABLE-INCOME * STATE-TAX-RATE

           *> FICA only up to wage base
           IF WS-GROSS-PAY <= FICA-WAGE-BASE
               COMPUTE WS-FICA-TAX ROUNDED =
                   WS-GROSS-PAY * FICA-RATE
           ELSE
               COMPUTE WS-FICA-TAX ROUNDED =
                   FICA-WAGE-BASE * FICA-RATE
           END-IF.

       2400-CALCULATE-DEDUCTIONS.
           COMPUTE WS-401K-DEDUCT ROUNDED =
               WS-GROSS-PAY * (EMP-401K-PCT / 100)

           EVALUATE TRUE
               WHEN HEALTH-SINGLE
                   MOVE HEALTH-SINGLE-AMT TO WS-HEALTH-DEDUCT
               WHEN HEALTH-FAMILY
                   MOVE HEALTH-FAMILY-AMT TO WS-HEALTH-DEDUCT
               WHEN HEALTH-NONE
                   MOVE ZEROS TO WS-HEALTH-DEDUCT
           END-EVALUATE.

       2500-CALCULATE-NET-PAY.
           COMPUTE WS-TOTAL-DEDUCT =
               WS-FEDERAL-TAX + WS-STATE-TAX + WS-FICA-TAX +
               WS-401K-DEDUCT + WS-HEALTH-DEDUCT
           COMPUTE WS-NET-PAY =
               WS-GROSS-PAY - WS-TOTAL-DEDUCT.

       2600-UPDATE-TOTALS.
           ADD 1                TO TOT-EMPLOYEES
           ADD WS-GROSS-PAY     TO TOT-GROSS-PAY
           ADD WS-FEDERAL-TAX   TO TOT-FEDERAL-TAX
           ADD WS-STATE-TAX     TO TOT-STATE-TAX
           ADD WS-FICA-TAX      TO TOT-FICA-TAX
           ADD WS-401K-DEDUCT   TO TOT-401K
           ADD WS-NET-PAY       TO TOT-NET-PAY

           *> Update department totals
           PERFORM VARYING DEPT-IDX FROM 1 BY 1
               UNTIL DEPT-IDX > 10
               IF DEPT-CODE(DEPT-IDX) = EMP-DEPT
                   ADD WS-GROSS-PAY TO DEPT-GROSS(DEPT-IDX)
                   ADD 1            TO DEPT-COUNT(DEPT-IDX)
               END-IF
           END-PERFORM.

       2700-PRINT-DETAIL.
           IF WS-LINE-COUNT >= WS-LINES-PER-PAGE
               ADD 1 TO WS-PAGE-NUM
               PERFORM 1200-PRINT-HEADERS
               MOVE ZEROS TO WS-LINE-COUNT
           END-IF

           MOVE EMP-ID TO RL-EMP-ID
           STRING FUNCTION TRIM(EMP-FIRST-NAME) ' '
                  FUNCTION TRIM(EMP-LAST-NAME)
               DELIMITED SIZE INTO RL-EMP-NAME
           MOVE WS-GROSS-PAY    TO RL-GROSS
           MOVE WS-FEDERAL-TAX  TO RL-FED-TAX
           MOVE WS-STATE-TAX    TO RL-STATE-TAX
           MOVE WS-FICA-TAX     TO RL-FICA
           MOVE WS-401K-DEDUCT  TO RL-401K
           MOVE WS-HEALTH-DEDUCT TO RL-HEALTH
           MOVE WS-NET-PAY      TO RL-NET-PAY

           WRITE REPORT-LINE FROM RL-DETAIL
           ADD 1 TO WS-LINE-COUNT.

       3000-PRINT-TOTALS.
           MOVE ALL '=' TO REPORT-LINE
           WRITE REPORT-LINE
           WRITE REPORT-LINE FROM SPACES

           *> Grand totals line
           MOVE SPACES TO RL-DETAIL
           MOVE ZEROS  TO RL-EMP-ID
           MOVE 'GRAND TOTALS' TO RL-EMP-NAME
           MOVE TOT-GROSS-PAY   TO RL-GROSS
           MOVE TOT-FEDERAL-TAX TO RL-FED-TAX
           MOVE TOT-STATE-TAX   TO RL-STATE-TAX
           MOVE TOT-FICA-TAX    TO RL-FICA
           MOVE TOT-401K        TO RL-401K
           MOVE TOT-NET-PAY     TO RL-NET-PAY
           WRITE REPORT-LINE FROM RL-DETAIL

           WRITE REPORT-LINE FROM SPACES
           MOVE ALL '-' TO REPORT-LINE
           WRITE REPORT-LINE

           *> Department summary
           MOVE 'DEPARTMENT SUMMARY:' TO REPORT-LINE
           WRITE REPORT-LINE
           PERFORM VARYING DEPT-IDX FROM 1 BY 1
               UNTIL DEPT-IDX > 10
               IF DEPT-COUNT(DEPT-IDX) > ZEROS
                   MOVE SPACES TO REPORT-LINE
                   STRING '  ' DEPT-NAME(DEPT-IDX)
                          ': ' DEPT-COUNT(DEPT-IDX)
                          ' employees, gross: '
                          DEPT-GROSS(DEPT-IDX)
                       DELIMITED SIZE INTO REPORT-LINE
                   WRITE REPORT-LINE
               END-IF
           END-PERFORM

           WRITE REPORT-LINE FROM SPACES
           MOVE SPACES TO REPORT-LINE
           STRING 'TOTAL EMPLOYEES PROCESSED: ' TOT-EMPLOYEES
               DELIMITED SIZE INTO REPORT-LINE
           WRITE REPORT-LINE.

       9000-TERMINATE.
           CLOSE EMPLOYEE-FILE
           CLOSE PAYROLL-REPORT
           DISPLAY "Payroll processing complete. "
                   TOT-EMPLOYEES " employees processed."
           DISPLAY "Report written to payroll_report.txt".
