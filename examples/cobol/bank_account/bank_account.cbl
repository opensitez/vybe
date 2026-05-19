      *> ============================================================
      *> BANK ACCOUNT TRANSACTION PROCESSING
      *> ============================================================
      *> Processes a stream of banking transactions against accounts.
      *> Handles deposits, withdrawals, transfers, interest posting.
      *> Produces account statements and exception reports.
      *>
      *> Demonstrates: INDEXED files, REWRITE, DELETE, dynamic
      *> access, SEARCH ALL (binary search), nested programs,
      *> EXCEPTION handling, COMPUTE with ON SIZE ERROR.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. BANK-ACCOUNT-SYSTEM.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT ACCOUNT-MASTER ASSIGN TO "accounts.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS ACCT-NUMBER
               ALTERNATE RECORD KEY IS ACCT-CUSTOMER-ID
                   WITH DUPLICATES
               FILE STATUS IS WS-ACCT-STATUS.

           SELECT TRANSACTION-FILE ASSIGN TO "transactions.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               ACCESS MODE IS SEQUENTIAL
               FILE STATUS IS WS-TXN-STATUS.

           SELECT STATEMENT-FILE ASSIGN TO "statements.txt"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-STMT-STATUS.

           SELECT EXCEPTION-LOG ASSIGN TO "exceptions.log"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-EXC-STATUS.

       DATA DIVISION.
       FILE SECTION.

       FD  ACCOUNT-MASTER
           RECORD CONTAINS 200 CHARACTERS.
       01  ACCOUNT-RECORD.
           05  ACCT-NUMBER         PIC X(12).
           05  ACCT-CUSTOMER-ID    PIC X(10).
           05  ACCT-CUSTOMER-NAME  PIC X(40).
           05  ACCT-TYPE           PIC X(2).
               88  CHECKING        VALUE 'CK'.
               88  SAVINGS         VALUE 'SV'.
               88  MONEY-MARKET    VALUE 'MM'.
               88  CD-ACCOUNT      VALUE 'CD'.
           05  ACCT-STATUS         PIC X(1).
               88  ACCT-ACTIVE     VALUE 'A'.
               88  ACCT-FROZEN     VALUE 'F'.
               88  ACCT-CLOSED     VALUE 'C'.
           05  ACCT-BALANCE        PIC S9(11)V99 COMP-3.
           05  ACCT-AVAILABLE-BAL  PIC S9(11)V99 COMP-3.
           05  ACCT-OPEN-DATE      PIC 9(8).
           05  ACCT-LAST-TXN-DATE  PIC 9(8).
           05  ACCT-INTEREST-RATE  PIC 9(2)V9(4) COMP-3.
           05  ACCT-OVERDRAFT-LIMIT PIC S9(7)V99 COMP-3.
           05  ACCT-MONTHLY-FEE    PIC 9(5)V99 COMP-3.
           05  ACCT-TXN-COUNT      PIC 9(8) COMP.
           05  ACCT-YTD-INTEREST   PIC S9(9)V99 COMP-3.
           05  FILLER              PIC X(47).

       FD  TRANSACTION-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  TRANSACTION-RECORD.
           05  TXN-ID              PIC X(12).
           05  TXN-ACCOUNT         PIC X(12).
           05  TXN-TYPE            PIC X(3).
               88  TXN-DEPOSIT     VALUE 'DEP'.
               88  TXN-WITHDRAWAL  VALUE 'WDR'.
               88  TXN-TRANSFER-OUT VALUE 'TFO'.
               88  TXN-TRANSFER-IN  VALUE 'TFI'.
               88  TXN-INTEREST    VALUE 'INT'.
               88  TXN-FEE         VALUE 'FEE'.
               88  TXN-REVERSAL    VALUE 'REV'.
           05  TXN-AMOUNT          PIC S9(9)V99.
           05  TXN-DATE            PIC 9(8).
           05  TXN-DESCRIPTION     PIC X(40).
           05  TXN-REFERENCE       PIC X(12).
           05  FILLER              PIC X(9).

       FD  STATEMENT-FILE
           RECORD CONTAINS 132 CHARACTERS.
       01  STATEMENT-LINE          PIC X(132).

       FD  EXCEPTION-LOG
           RECORD CONTAINS 200 CHARACTERS.
       01  EXCEPTION-RECORD        PIC X(200).

       WORKING-STORAGE SECTION.

       01  WS-STATUS-CODES.
           05  WS-ACCT-STATUS      PIC XX.
               88  ACCT-OK         VALUE '00'.
               88  ACCT-NOT-FOUND  VALUE '23'.
               88  ACCT-DUP-KEY    VALUE '22'.
               88  ACCT-EOF        VALUE '10'.
           05  WS-TXN-STATUS       PIC XX.
               88  TXN-OK          VALUE '00'.
               88  TXN-EOF         VALUE '10'.
           05  WS-STMT-STATUS      PIC XX.
           05  WS-EXC-STATUS       PIC XX.

       01  WS-WORK-FIELDS.
           05  WS-NEW-BALANCE      PIC S9(11)V99.
           05  WS-INTEREST-AMOUNT  PIC S9(9)V99.
           05  WS-FEE-AMOUNT       PIC S9(5)V99.
           05  WS-TRANSFER-ACCT    PIC X(12).
           05  WS-CURRENT-DATE     PIC 9(8).
           05  WS-PROCESS-DATE     PIC 9(8).
           05  WS-ERROR-MSG        PIC X(80).
           05  WS-EXCEPTION-LINE   PIC X(200).

       01  WS-COUNTERS.
           05  WS-TXN-PROCESSED    PIC 9(8) VALUE ZEROS.
           05  WS-TXN-ACCEPTED     PIC 9(8) VALUE ZEROS.
           05  WS-TXN-REJECTED     PIC 9(8) VALUE ZEROS.
           05  WS-OVERDRAFTS       PIC 9(6) VALUE ZEROS.
           05  WS-TOTAL-DEPOSITS   PIC S9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-WITHDRAWALS PIC S9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-FEES       PIC S9(11)V99 VALUE ZEROS.
           05  WS-TOTAL-INTEREST   PIC S9(11)V99 VALUE ZEROS.

       01  WS-STATEMENT-FIELDS.
           05  WS-STMT-ACCT        PIC X(12) VALUE SPACES.
           05  WS-STMT-LINE-COUNT  PIC 9(4)  VALUE ZEROS.

       01  WS-FORMATTED.
           05  WF-BALANCE          PIC -ZZZ,ZZZ,ZZ9.99.
           05  WF-AMOUNT           PIC -ZZZ,ZZZ,ZZ9.99.
           05  WF-DATE             PIC 9999/99/99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-OPEN-FILES
           PERFORM 2000-PROCESS-TRANSACTIONS
               UNTIL TXN-EOF
           PERFORM 3000-POST-MONTHLY-FEES
           PERFORM 4000-PRINT-SUMMARY
           PERFORM 9000-CLOSE-FILES
           STOP RUN.

       1000-OPEN-FILES.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           OPEN I-O    ACCOUNT-MASTER
           OPEN INPUT  TRANSACTION-FILE
           OPEN OUTPUT STATEMENT-FILE
           OPEN OUTPUT EXCEPTION-LOG
           PERFORM 1100-READ-TRANSACTION.

       1100-READ-TRANSACTION.
           READ TRANSACTION-FILE
               AT END MOVE '10' TO WS-TXN-STATUS
           END-READ.

       2000-PROCESS-TRANSACTIONS.
           ADD 1 TO WS-TXN-PROCESSED
           PERFORM 2100-READ-ACCOUNT
           IF ACCT-OK
               PERFORM 2200-VALIDATE-TRANSACTION
               IF WS-ERROR-MSG = SPACES
                   PERFORM 2300-APPLY-TRANSACTION
                   PERFORM 2400-UPDATE-ACCOUNT
                   ADD 1 TO WS-TXN-ACCEPTED
               ELSE
                   PERFORM 2500-LOG-EXCEPTION
                   ADD 1 TO WS-TXN-REJECTED
               END-IF
           ELSE
               MOVE 'ACCOUNT NOT FOUND' TO WS-ERROR-MSG
               PERFORM 2500-LOG-EXCEPTION
               ADD 1 TO WS-TXN-REJECTED
           END-IF
           PERFORM 1100-READ-TRANSACTION.

       2100-READ-ACCOUNT.
           MOVE TXN-ACCOUNT TO ACCT-NUMBER
           READ ACCOUNT-MASTER
               INVALID KEY
                   MOVE '23' TO WS-ACCT-STATUS
           END-READ.

       2200-VALIDATE-TRANSACTION.
           MOVE SPACES TO WS-ERROR-MSG
           EVALUATE TRUE
               WHEN ACCT-FROZEN
                   MOVE 'ACCOUNT IS FROZEN' TO WS-ERROR-MSG
               WHEN ACCT-CLOSED
                   MOVE 'ACCOUNT IS CLOSED' TO WS-ERROR-MSG
               WHEN TXN-WITHDRAWAL OR TXN-TRANSFER-OUT
                   COMPUTE WS-NEW-BALANCE =
                       ACCT-AVAILABLE-BAL - TXN-AMOUNT
                   IF WS-NEW-BALANCE < ACCT-OVERDRAFT-LIMIT * -1
                       MOVE 'INSUFFICIENT FUNDS' TO WS-ERROR-MSG
                       ADD 1 TO WS-OVERDRAFTS
                   END-IF
               WHEN TXN-AMOUNT < ZEROS
                   MOVE 'NEGATIVE AMOUNT NOT ALLOWED' TO WS-ERROR-MSG
               WHEN OTHER
                   CONTINUE
           END-EVALUATE.

       2300-APPLY-TRANSACTION.
           EVALUATE TRUE
               WHEN TXN-DEPOSIT OR TXN-TRANSFER-IN OR TXN-INTEREST
                   ADD TXN-AMOUNT TO ACCT-BALANCE
                   ADD TXN-AMOUNT TO ACCT-AVAILABLE-BAL
                   IF TXN-DEPOSIT
                       ADD TXN-AMOUNT TO WS-TOTAL-DEPOSITS
                   END-IF
                   IF TXN-INTEREST
                       ADD TXN-AMOUNT TO WS-TOTAL-INTEREST
                       ADD TXN-AMOUNT TO ACCT-YTD-INTEREST
                   END-IF
               WHEN TXN-WITHDRAWAL OR TXN-TRANSFER-OUT
                   SUBTRACT TXN-AMOUNT FROM ACCT-BALANCE
                   SUBTRACT TXN-AMOUNT FROM ACCT-AVAILABLE-BAL
                   ADD TXN-AMOUNT TO WS-TOTAL-WITHDRAWALS
               WHEN TXN-FEE
                   SUBTRACT TXN-AMOUNT FROM ACCT-BALANCE
                   SUBTRACT TXN-AMOUNT FROM ACCT-AVAILABLE-BAL
                   ADD TXN-AMOUNT TO WS-TOTAL-FEES
               WHEN TXN-REVERSAL
                   *> Reverse the original transaction amount
                   ADD TXN-AMOUNT TO ACCT-BALANCE
                   ADD TXN-AMOUNT TO ACCT-AVAILABLE-BAL
           END-EVALUATE
           ADD 1 TO ACCT-TXN-COUNT
           MOVE WS-CURRENT-DATE TO ACCT-LAST-TXN-DATE.

       2400-UPDATE-ACCOUNT.
           REWRITE ACCOUNT-RECORD
               INVALID KEY
                   STRING 'REWRITE FAILED FOR ACCOUNT: '
                          ACCT-NUMBER
                       DELIMITED SIZE INTO WS-ERROR-MSG
                   PERFORM 2500-LOG-EXCEPTION
           END-REWRITE.

       2500-LOG-EXCEPTION.
           STRING TXN-ID        ' | '
                  TXN-ACCOUNT   ' | '
                  TXN-TYPE      ' | '
                  TXN-AMOUNT    ' | '
                  WS-ERROR-MSG
               DELIMITED SIZE INTO EXCEPTION-RECORD
           WRITE EXCEPTION-RECORD.

       3000-POST-MONTHLY-FEES.
           *> Scan all accounts and post monthly maintenance fees
           MOVE LOW-VALUES TO ACCT-NUMBER
           START ACCOUNT-MASTER KEY >= ACCT-NUMBER
               INVALID KEY STOP RUN
           END-START
           PERFORM 3100-FEE-LOOP
               UNTIL ACCT-EOF.

       3100-FEE-LOOP.
           READ ACCOUNT-MASTER NEXT
               AT END MOVE '10' TO WS-ACCT-STATUS
           END-READ
           IF NOT ACCT-EOF
               IF ACCT-ACTIVE AND ACCT-MONTHLY-FEE > ZEROS
                   SUBTRACT ACCT-MONTHLY-FEE FROM ACCT-BALANCE
                   SUBTRACT ACCT-MONTHLY-FEE FROM ACCT-AVAILABLE-BAL
                   ADD ACCT-MONTHLY-FEE TO WS-TOTAL-FEES
                   REWRITE ACCOUNT-RECORD
                       INVALID KEY CONTINUE
                   END-REWRITE
               END-IF
           END-IF.

       4000-PRINT-SUMMARY.
           MOVE WS-TXN-PROCESSED   TO WF-AMOUNT
           DISPLAY "=== TRANSACTION PROCESSING SUMMARY ==="
           DISPLAY "Transactions Processed : " WS-TXN-PROCESSED
           DISPLAY "Transactions Accepted  : " WS-TXN-ACCEPTED
           DISPLAY "Transactions Rejected  : " WS-TXN-REJECTED
           DISPLAY "Overdraft Attempts     : " WS-OVERDRAFTS
           MOVE WS-TOTAL-DEPOSITS TO WF-BALANCE
           DISPLAY "Total Deposits         : " WF-BALANCE
           MOVE WS-TOTAL-WITHDRAWALS TO WF-BALANCE
           DISPLAY "Total Withdrawals      : " WF-BALANCE
           MOVE WS-TOTAL-FEES TO WF-BALANCE
           DISPLAY "Total Fees Collected   : " WF-BALANCE
           MOVE WS-TOTAL-INTEREST TO WF-BALANCE
           DISPLAY "Total Interest Posted  : " WF-BALANCE.

       9000-CLOSE-FILES.
           CLOSE ACCOUNT-MASTER
                 TRANSACTION-FILE
                 STATEMENT-FILE
                 EXCEPTION-LOG.
