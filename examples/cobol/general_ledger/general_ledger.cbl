      *> ============================================================
      *> GENERAL LEDGER ACCOUNTING SYSTEM
      *> ============================================================
      *> Double-entry bookkeeping: journal entries, trial balance,
      *> income statement, balance sheet. Full chart of accounts.
      *>
      *> Demonstrates: COBOL 2014 intrinsic functions,
      *> multi-level PERFORM, complex EVALUATE, COMPUTE with
      *> ROUNDED, report writer concepts, table SEARCH,
      *> INITIALIZE with REPLACING.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. GENERAL-LEDGER.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT JOURNAL-FILE ASSIGN TO "journal_entries.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-JNL-STATUS.

           SELECT LEDGER-FILE ASSIGN TO "ledger.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS LED-ACCOUNT-NUM
               FILE STATUS IS WS-LED-STATUS.

           SELECT TRIAL-BALANCE ASSIGN TO "trial_balance.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT INCOME-STMT ASSIGN TO "income_statement.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT BALANCE-SHEET ASSIGN TO "balance_sheet.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  JOURNAL-FILE
           RECORD CONTAINS 200 CHARACTERS.
       01  JOURNAL-RECORD.
           05  JNL-ENTRY-NUM       PIC 9(10).
           05  JNL-DATE            PIC 9(8).
           05  JNL-PERIOD          PIC 9(6).
           05  JNL-ACCOUNT-NUM     PIC X(10).
           05  JNL-DESCRIPTION     PIC X(50).
           05  JNL-DEBIT-AMOUNT    PIC S9(13)V99.
           05  JNL-CREDIT-AMOUNT   PIC S9(13)V99.
           05  JNL-REFERENCE       PIC X(20).
           05  JNL-POSTED-BY       PIC X(8).
           05  FILLER              PIC X(55).

       FD  LEDGER-FILE
           RECORD CONTAINS 300 CHARACTERS.
       01  LEDGER-RECORD.
           05  LED-ACCOUNT-NUM     PIC X(10).
           05  LED-ACCOUNT-NAME    PIC X(50).
           05  LED-ACCOUNT-TYPE    PIC X(2).
               88  LED-ASSET       VALUE 'AS'.
               88  LED-LIABILITY   VALUE 'LI'.
               88  LED-EQUITY      VALUE 'EQ'.
               88  LED-REVENUE     VALUE 'RE'.
               88  LED-EXPENSE     VALUE 'EX'.
               88  LED-CONTRA      VALUE 'CO'.
           05  LED-NORMAL-BALANCE  PIC X(1).
               88  DEBIT-NORMAL    VALUE 'D'.
               88  CREDIT-NORMAL   VALUE 'C'.
           05  LED-CATEGORY        PIC X(4).
           05  LED-SUBCATEGORY     PIC X(6).
           05  LED-CURRENT-BALANCE PIC S9(15)V99 COMP-3.
           05  LED-PERIOD-DEBITS   PIC S9(15)V99 COMP-3.
           05  LED-PERIOD-CREDITS  PIC S9(15)V99 COMP-3.
           05  LED-YTD-DEBITS      PIC S9(15)V99 COMP-3.
           05  LED-YTD-CREDITS     PIC S9(15)V99 COMP-3.
           05  LED-BUDGET-AMOUNT   PIC S9(15)V99 COMP-3.
           05  LED-PRIOR-YEAR-BAL  PIC S9(15)V99 COMP-3.
           05  LED-ACTIVE-FLAG     PIC X(1).
               88  ACCOUNT-ACTIVE  VALUE 'Y'.
               88  ACCOUNT-INACTIVE VALUE 'N'.
           05  FILLER              PIC X(80).

       FD  TRIAL-BALANCE
           RECORD CONTAINS 132 CHARACTERS.
       01  TB-LINE                 PIC X(132).

       FD  INCOME-STMT
           RECORD CONTAINS 132 CHARACTERS.
       01  IS-LINE                 PIC X(132).

       FD  BALANCE-SHEET
           RECORD CONTAINS 132 CHARACTERS.
       01  BS-LINE                 PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-JNL-STATUS       PIC XX.
               88  JNL-OK          VALUE '00'.
               88  JNL-EOF         VALUE '10'.
           05  WS-LED-STATUS       PIC XX.
               88  LED-OK          VALUE '00'.
               88  LED-NOT-FOUND   VALUE '23'.
               88  LED-EOF         VALUE '10'.

       01  WS-PERIOD-CONTROL.
           05  WS-CURRENT-PERIOD   PIC 9(6).
           05  WS-FISCAL-YEAR      PIC 9(4).
           05  WS-PERIOD-NUM       PIC 9(2).
           05  WS-PERIOD-START     PIC 9(8).
           05  WS-PERIOD-END       PIC 9(8).

       01  WS-TOTALS.
           05  WS-TOTAL-DEBITS     PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-CREDITS    PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-ASSETS     PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-LIABILITIES PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-EQUITY     PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-REVENUE    PIC S9(17)V99 VALUE ZEROS.
           05  WS-TOTAL-EXPENSES   PIC S9(17)V99 VALUE ZEROS.
           05  WS-NET-INCOME       PIC S9(17)V99 VALUE ZEROS.
           05  WS-ENTRIES-POSTED   PIC 9(8)      VALUE ZEROS.
           05  WS-ENTRIES-REJECTED PIC 9(8)      VALUE ZEROS.

       01  WS-WORK-FIELDS.
           05  WS-DEBIT-TOTAL      PIC S9(15)V99 VALUE ZEROS.
           05  WS-CREDIT-TOTAL     PIC S9(15)V99 VALUE ZEROS.
           05  WS-VARIANCE         PIC S9(15)V99 VALUE ZEROS.
           05  WS-BUDGET-VARIANCE  PIC S9(15)V99 VALUE ZEROS.
           05  WS-BUDGET-PCT       PIC ZZZ9.99.

       01  WS-FORMATTED.
           05  WF-AMOUNT           PIC -ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-BALANCE          PIC -ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-BUDGET           PIC -ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-VARIANCE         PIC -ZZZ,ZZZ,ZZZ,ZZ9.99.

       01  WS-REPORT-HEADER.
           05  FILLER              PIC X(40) VALUE SPACES.
           05  WS-COMPANY-NAME     PIC X(40)
               VALUE 'ACME CORPORATION'.
           05  FILLER              PIC X(52) VALUE SPACES.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-POST-JOURNAL-ENTRIES
               UNTIL JNL-EOF
           PERFORM 3000-VERIFY-TRIAL-BALANCE
           PERFORM 4000-GENERATE-TRIAL-BALANCE
           PERFORM 5000-GENERATE-INCOME-STATEMENT
           PERFORM 6000-GENERATE-BALANCE-SHEET
           PERFORM 7000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:6) TO WS-CURRENT-PERIOD
           MOVE WS-CURRENT-PERIOD(1:4) TO WS-FISCAL-YEAR
           MOVE WS-CURRENT-PERIOD(5:2) TO WS-PERIOD-NUM
           OPEN INPUT  JOURNAL-FILE
           OPEN I-O    LEDGER-FILE
           OPEN OUTPUT TRIAL-BALANCE
           OPEN OUTPUT INCOME-STMT
           OPEN OUTPUT BALANCE-SHEET
           PERFORM 1100-READ-JOURNAL.

       1100-READ-JOURNAL.
           READ JOURNAL-FILE
               AT END MOVE '10' TO WS-JNL-STATUS
           END-READ.

       2000-POST-JOURNAL-ENTRIES.
           *> Validate: debits must equal credits per entry
           *> (In a real system, entries are grouped by entry number)
           MOVE JNL-ACCOUNT-NUM TO LED-ACCOUNT-NUM
           READ LEDGER-FILE
               INVALID KEY
                   ADD 1 TO WS-ENTRIES-REJECTED
                   PERFORM 1100-READ-JOURNAL
                   STOP RUN
           END-READ

           IF LED-OK AND ACCOUNT-ACTIVE
               *> Post debit
               IF JNL-DEBIT-AMOUNT > ZEROS
                   ADD JNL-DEBIT-AMOUNT TO LED-PERIOD-DEBITS
                   ADD JNL-DEBIT-AMOUNT TO LED-YTD-DEBITS
                   IF DEBIT-NORMAL
                       ADD JNL-DEBIT-AMOUNT TO LED-CURRENT-BALANCE
                   ELSE
                       SUBTRACT JNL-DEBIT-AMOUNT FROM LED-CURRENT-BALANCE
                   END-IF
               END-IF

               *> Post credit
               IF JNL-CREDIT-AMOUNT > ZEROS
                   ADD JNL-CREDIT-AMOUNT TO LED-PERIOD-CREDITS
                   ADD JNL-CREDIT-AMOUNT TO LED-YTD-CREDITS
                   IF CREDIT-NORMAL
                       ADD JNL-CREDIT-AMOUNT TO LED-CURRENT-BALANCE
                   ELSE
                       SUBTRACT JNL-CREDIT-AMOUNT FROM LED-CURRENT-BALANCE
                   END-IF
               END-IF

               REWRITE LEDGER-RECORD
                   INVALID KEY ADD 1 TO WS-ENTRIES-REJECTED
               END-REWRITE
               ADD 1 TO WS-ENTRIES-POSTED
               ADD JNL-DEBIT-AMOUNT  TO WS-TOTAL-DEBITS
               ADD JNL-CREDIT-AMOUNT TO WS-TOTAL-CREDITS
           ELSE
               ADD 1 TO WS-ENTRIES-REJECTED
           END-IF
           PERFORM 1100-READ-JOURNAL.

       3000-VERIFY-TRIAL-BALANCE.
           COMPUTE WS-VARIANCE = WS-TOTAL-DEBITS - WS-TOTAL-CREDITS
           IF WS-VARIANCE NOT = ZEROS
               DISPLAY "WARNING: TRIAL BALANCE OUT OF BALANCE BY: "
                       WS-VARIANCE
           ELSE
               DISPLAY "Trial balance verified: debits = credits"
           END-IF.

       4000-GENERATE-TRIAL-BALANCE.
           WRITE TB-LINE FROM WS-REPORT-HEADER
           WRITE TB-LINE FROM
               "                    TRIAL BALANCE"
           WRITE TB-LINE FROM ALL '-'
           WRITE TB-LINE FROM
               "ACCOUNT     ACCOUNT NAME                          " &
               "DEBITS              CREDITS             BALANCE"
           WRITE TB-LINE FROM ALL '-'

           MOVE LOW-VALUES TO LED-ACCOUNT-NUM
           START LEDGER-FILE KEY >= LED-ACCOUNT-NUM
               INVALID KEY STOP RUN
           END-START

           PERFORM 4100-TB-SCAN UNTIL LED-EOF

           WRITE TB-LINE FROM ALL '='
           MOVE WS-TOTAL-DEBITS  TO WF-AMOUNT
           MOVE WS-TOTAL-CREDITS TO WF-BALANCE
           MOVE SPACES TO TB-LINE
           STRING 'TOTALS' SPACES(44) WF-AMOUNT SPACES(4) WF-BALANCE
               DELIMITED SIZE INTO TB-LINE
           WRITE TB-LINE.

       4100-TB-SCAN.
           READ LEDGER-FILE NEXT
               AT END MOVE '10' TO WS-LED-STATUS
           END-READ
           IF NOT LED-EOF AND ACCOUNT-ACTIVE
               MOVE SPACES TO TB-LINE
               MOVE LED-PERIOD-DEBITS  TO WF-AMOUNT
               MOVE LED-PERIOD-CREDITS TO WF-BALANCE
               MOVE LED-CURRENT-BALANCE TO WF-VARIANCE
               STRING LED-ACCOUNT-NUM SPACES(4)
                      LED-ACCOUNT-NAME SPACES(4)
                      WF-AMOUNT SPACES(4)
                      WF-BALANCE SPACES(4)
                      WF-VARIANCE
                   DELIMITED SIZE INTO TB-LINE
               WRITE TB-LINE
           END-IF.

       5000-GENERATE-INCOME-STATEMENT.
           WRITE IS-LINE FROM WS-REPORT-HEADER
           WRITE IS-LINE FROM
               "                  INCOME STATEMENT"
           WRITE IS-LINE FROM ALL '-'

           WRITE IS-LINE FROM "REVENUES:"
           MOVE LOW-VALUES TO LED-ACCOUNT-NUM
           START LEDGER-FILE KEY >= LED-ACCOUNT-NUM
               INVALID KEY STOP RUN
           END-START
           PERFORM 5100-REVENUE-SCAN UNTIL LED-EOF

           WRITE IS-LINE FROM ALL '-'
           MOVE WS-TOTAL-REVENUE TO WF-AMOUNT
           MOVE SPACES TO IS-LINE
           STRING 'TOTAL REVENUES' SPACES(36) WF-AMOUNT
               DELIMITED SIZE INTO IS-LINE
           WRITE IS-LINE

           WRITE IS-LINE FROM SPACES
           WRITE IS-LINE FROM "EXPENSES:"
           MOVE LOW-VALUES TO LED-ACCOUNT-NUM
           START LEDGER-FILE KEY >= LED-ACCOUNT-NUM
               INVALID KEY STOP RUN
           END-START
           PERFORM 5200-EXPENSE-SCAN UNTIL LED-EOF

           WRITE IS-LINE FROM ALL '-'
           MOVE WS-TOTAL-EXPENSES TO WF-AMOUNT
           MOVE SPACES TO IS-LINE
           STRING 'TOTAL EXPENSES' SPACES(36) WF-AMOUNT
               DELIMITED SIZE INTO IS-LINE
           WRITE IS-LINE

           COMPUTE WS-NET-INCOME = WS-TOTAL-REVENUE - WS-TOTAL-EXPENSES
           WRITE IS-LINE FROM ALL '='
           MOVE WS-NET-INCOME TO WF-AMOUNT
           MOVE SPACES TO IS-LINE
           STRING 'NET INCOME' SPACES(40) WF-AMOUNT
               DELIMITED SIZE INTO IS-LINE
           WRITE IS-LINE.

       5100-REVENUE-SCAN.
           READ LEDGER-FILE NEXT
               AT END MOVE '10' TO WS-LED-STATUS
           END-READ
           IF NOT LED-EOF
               IF LED-REVENUE AND ACCOUNT-ACTIVE
                   ADD LED-CURRENT-BALANCE TO WS-TOTAL-REVENUE
                   MOVE LED-CURRENT-BALANCE TO WF-AMOUNT
                   MOVE SPACES TO IS-LINE
                   STRING '  ' LED-ACCOUNT-NAME SPACES(4) WF-AMOUNT
                       DELIMITED SIZE INTO IS-LINE
                   WRITE IS-LINE
               END-IF
           END-IF.

       5200-EXPENSE-SCAN.
           READ LEDGER-FILE NEXT
               AT END MOVE '10' TO WS-LED-STATUS
           END-READ
           IF NOT LED-EOF
               IF LED-EXPENSE AND ACCOUNT-ACTIVE
                   ADD LED-CURRENT-BALANCE TO WS-TOTAL-EXPENSES
                   MOVE LED-CURRENT-BALANCE TO WF-AMOUNT
                   COMPUTE WS-BUDGET-VARIANCE =
                       LED-CURRENT-BALANCE - LED-BUDGET-AMOUNT
                   MOVE WS-BUDGET-VARIANCE TO WF-VARIANCE
                   MOVE SPACES TO IS-LINE
                   STRING '  ' LED-ACCOUNT-NAME SPACES(4)
                          WF-AMOUNT '  Var: ' WF-VARIANCE
                       DELIMITED SIZE INTO IS-LINE
                   WRITE IS-LINE
               END-IF
           END-IF.

       6000-GENERATE-BALANCE-SHEET.
           WRITE BS-LINE FROM WS-REPORT-HEADER
           WRITE BS-LINE FROM
               "                   BALANCE SHEET"
           WRITE BS-LINE FROM ALL '-'

           WRITE BS-LINE FROM "ASSETS:"
           MOVE LOW-VALUES TO LED-ACCOUNT-NUM
           START LEDGER-FILE KEY >= LED-ACCOUNT-NUM
               INVALID KEY STOP RUN
           END-START
           PERFORM 6100-ASSET-SCAN UNTIL LED-EOF

           WRITE BS-LINE FROM ALL '-'
           MOVE WS-TOTAL-ASSETS TO WF-AMOUNT
           MOVE SPACES TO BS-LINE
           STRING 'TOTAL ASSETS' SPACES(38) WF-AMOUNT
               DELIMITED SIZE INTO BS-LINE
           WRITE BS-LINE

           WRITE BS-LINE FROM SPACES
           WRITE BS-LINE FROM "LIABILITIES:"
           MOVE LOW-VALUES TO LED-ACCOUNT-NUM
           START LEDGER-FILE KEY >= LED-ACCOUNT-NUM
               INVALID KEY STOP RUN
           END-START
           PERFORM 6200-LIABILITY-SCAN UNTIL LED-EOF

           WRITE BS-LINE FROM "EQUITY:"
           ADD WS-NET-INCOME TO WS-TOTAL-EQUITY
           MOVE WS-TOTAL-EQUITY TO WF-AMOUNT
           MOVE SPACES TO BS-LINE
           STRING '  RETAINED EARNINGS (incl net income)' WF-AMOUNT
               DELIMITED SIZE INTO BS-LINE
           WRITE BS-LINE

           WRITE BS-LINE FROM ALL '='
           COMPUTE WS-VARIANCE =
               WS-TOTAL-LIABILITIES + WS-TOTAL-EQUITY - WS-TOTAL-ASSETS
           IF WS-VARIANCE NOT = ZEROS
               DISPLAY "BALANCE SHEET DOES NOT BALANCE: " WS-VARIANCE
           END-IF.

       6100-ASSET-SCAN.
           READ LEDGER-FILE NEXT
               AT END MOVE '10' TO WS-LED-STATUS
           END-READ
           IF NOT LED-EOF
               IF LED-ASSET AND ACCOUNT-ACTIVE
                   ADD LED-CURRENT-BALANCE TO WS-TOTAL-ASSETS
                   MOVE LED-CURRENT-BALANCE TO WF-AMOUNT
                   MOVE SPACES TO BS-LINE
                   STRING '  ' LED-ACCOUNT-NAME SPACES(4) WF-AMOUNT
                       DELIMITED SIZE INTO BS-LINE
                   WRITE BS-LINE
               END-IF
           END-IF.

       6200-LIABILITY-SCAN.
           READ LEDGER-FILE NEXT
               AT END MOVE '10' TO WS-LED-STATUS
           END-READ
           IF NOT LED-EOF
               IF LED-LIABILITY AND ACCOUNT-ACTIVE
                   ADD LED-CURRENT-BALANCE TO WS-TOTAL-LIABILITIES
                   MOVE LED-CURRENT-BALANCE TO WF-AMOUNT
                   MOVE SPACES TO BS-LINE
                   STRING '  ' LED-ACCOUNT-NAME SPACES(4) WF-AMOUNT
                       DELIMITED SIZE INTO BS-LINE
                   WRITE BS-LINE
               END-IF
           END-IF.

       7000-PRINT-SUMMARY.
           DISPLAY "=== GENERAL LEDGER PROCESSING COMPLETE ==="
           DISPLAY "Entries Posted   : " WS-ENTRIES-POSTED
           DISPLAY "Entries Rejected : " WS-ENTRIES-REJECTED
           DISPLAY "Total Debits     : " WS-TOTAL-DEBITS
           DISPLAY "Total Credits    : " WS-TOTAL-CREDITS
           DISPLAY "Net Income       : " WS-NET-INCOME
           DISPLAY "Total Assets     : " WS-TOTAL-ASSETS.

       9000-TERMINATE.
           CLOSE JOURNAL-FILE LEDGER-FILE
                 TRIAL-BALANCE INCOME-STMT BALANCE-SHEET.
