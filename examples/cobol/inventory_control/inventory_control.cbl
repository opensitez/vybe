      *> ============================================================
      *> INVENTORY CONTROL SYSTEM
      *> ============================================================
      *> Full warehouse inventory management: stock receipts,
      *> issues, adjustments, reorder processing, valuation.
      *> Uses FIFO costing for inventory valuation.
      *>
      *> Demonstrates: RELATIVE files, OCCURS DEPENDING ON,
      *> INSPECT, MOVE CORRESPONDING, multi-level tables,
      *> SORT verb, MERGE, CALL to subprogram.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INVENTORY-CONTROL.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT INVENTORY-MASTER ASSIGN TO "inventory.rel"
               ORGANIZATION IS RELATIVE
               ACCESS MODE IS DYNAMIC
               RELATIVE KEY IS WS-REL-KEY
               FILE STATUS IS WS-INV-STATUS.

           SELECT TRANSACTION-FILE ASSIGN TO "inv_transactions.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-TXN-STATUS.

           SELECT SORT-WORK-FILE ASSIGN TO "sort_work.tmp"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT REORDER-REPORT ASSIGN TO "reorder_report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT VALUATION-REPORT ASSIGN TO "valuation_report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  INVENTORY-MASTER
           RECORD CONTAINS 500 CHARACTERS.
       01  INVENTORY-RECORD.
           05  INV-ITEM-CODE       PIC X(12).
           05  INV-DESCRIPTION     PIC X(50).
           05  INV-CATEGORY        PIC X(6).
           05  INV-UNIT-OF-MEASURE PIC X(4).
           05  INV-LOCATION.
               10  INV-WAREHOUSE   PIC X(4).
               10  INV-AISLE       PIC X(3).
               10  INV-BIN         PIC X(4).
           05  INV-QTY-ON-HAND     PIC S9(9)V999 COMP-3.
           05  INV-QTY-RESERVED    PIC S9(9)V999 COMP-3.
           05  INV-QTY-ON-ORDER    PIC S9(9)V999 COMP-3.
           05  INV-REORDER-POINT   PIC 9(9)V999  COMP-3.
           05  INV-REORDER-QTY     PIC 9(9)V999  COMP-3.
           05  INV-MAX-STOCK       PIC 9(9)V999  COMP-3.
           05  INV-LEAD-TIME-DAYS  PIC 9(3).
           05  INV-LAST-RECEIPT    PIC 9(8).
           05  INV-LAST-ISSUE      PIC 9(8).
           05  INV-SUPPLIER-CODE   PIC X(8).
           05  INV-FIFO-LAYERS     PIC 9(2).
           05  INV-FIFO-TABLE.
               10  INV-FIFO-ENTRY OCCURS 1 TO 20 TIMES
                   DEPENDING ON INV-FIFO-LAYERS.
                   15  FIFO-RECEIPT-DATE PIC 9(8).
                   15  FIFO-QTY          PIC 9(9)V999 COMP-3.
                   15  FIFO-UNIT-COST    PIC 9(9)V9999 COMP-3.
           05  INV-AVG-COST        PIC 9(9)V9999 COMP-3.
           05  INV-TOTAL-VALUE     PIC S9(13)V99 COMP-3.
           05  INV-YTD-ISSUES      PIC 9(11)V999 COMP-3.
           05  INV-YTD-RECEIPTS    PIC 9(11)V999 COMP-3.
           05  FILLER              PIC X(50).

       FD  TRANSACTION-FILE
           RECORD CONTAINS 120 CHARACTERS.
       01  TXN-RECORD.
           05  TXN-SEQ-NUM         PIC 9(8).
           05  TXN-DATE            PIC 9(8).
           05  TXN-TYPE            PIC X(3).
               88  TXN-RECEIPT     VALUE 'REC'.
               88  TXN-ISSUE       VALUE 'ISS'.
               88  TXN-ADJUST-UP   VALUE 'ADJ'.
               88  TXN-ADJUST-DOWN VALUE 'ADJ'.
               88  TXN-TRANSFER    VALUE 'TRF'.
               88  TXN-RETURN      VALUE 'RET'.
               88  TXN-SCRAP       VALUE 'SCR'.
           05  TXN-ITEM-CODE       PIC X(12).
           05  TXN-QUANTITY        PIC S9(9)V999.
           05  TXN-UNIT-COST       PIC 9(9)V9999.
           05  TXN-REFERENCE       PIC X(20).
           05  TXN-WAREHOUSE       PIC X(4).
           05  TXN-REASON-CODE     PIC X(4).
           05  FILLER              PIC X(40).

       SD  SORT-WORK-FILE
           RECORD CONTAINS 120 CHARACTERS.
       01  SORT-RECORD             PIC X(120).

       FD  REORDER-REPORT
           RECORD CONTAINS 132 CHARACTERS.
       01  REORDER-LINE            PIC X(132).

       FD  VALUATION-REPORT
           RECORD CONTAINS 132 CHARACTERS.
       01  VALUATION-LINE          PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-INV-STATUS       PIC XX.
               88  INV-OK          VALUE '00'.
               88  INV-NOT-FOUND   VALUE '23'.
               88  INV-EOF         VALUE '10'.
           05  WS-TXN-STATUS       PIC XX.
               88  TXN-OK          VALUE '00'.
               88  TXN-EOF         VALUE '10'.
           05  WS-REL-KEY          PIC 9(8).

       01  WS-WORK-AREA.
           05  WS-ITEM-HASH        PIC 9(8).
           05  WS-FIFO-VALUE       PIC S9(13)V99.
           05  WS-ISSUE-QTY-REMAIN PIC 9(9)V999.
           05  WS-ISSUE-COST       PIC S9(13)V99.
           05  WS-LAYER-IDX        PIC 9(2).
           05  WS-CURRENT-DATE     PIC 9(8).
           05  WS-ERROR-FLAG       PIC X VALUE 'N'.
               88  HAS-ERROR       VALUE 'Y'.
               88  NO-ERROR        VALUE 'N'.

       01  WS-COUNTERS.
           05  WS-RECEIPTS-COUNT   PIC 9(8) VALUE ZEROS.
           05  WS-ISSUES-COUNT     PIC 9(8) VALUE ZEROS.
           05  WS-ADJUSTMENTS      PIC 9(8) VALUE ZEROS.
           05  WS-ERRORS           PIC 9(8) VALUE ZEROS.
           05  WS-REORDER-COUNT    PIC 9(6) VALUE ZEROS.
           05  WS-TOTAL-INV-VALUE  PIC S9(15)V99 VALUE ZEROS.

       01  WS-CATEGORY-TABLE.
           05  CAT-ENTRY OCCURS 20 TIMES INDEXED BY CAT-IDX.
               10  CAT-CODE        PIC X(6).
               10  CAT-DESC        PIC X(30).
               10  CAT-VALUE       PIC S9(13)V99 VALUE ZEROS.
               10  CAT-ITEM-COUNT  PIC 9(6)      VALUE ZEROS.

       01  WS-REPORT-LINES.
           05  RL-REORDER.
               10  RL-RO-ITEM      PIC X(12).
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-DESC      PIC X(30).
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-SUPPLIER  PIC X(8).
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-ON-HAND   PIC ZZZ,ZZZ,ZZ9.999.
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-REORDER   PIC ZZZ,ZZZ,ZZ9.999.
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-SUGGEST   PIC ZZZ,ZZZ,ZZ9.999.
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-RO-LEAD      PIC ZZ9.
               10  FILLER          PIC X(15) VALUE SPACES.
           05  RL-VALUATION.
               10  RL-VAL-ITEM     PIC X(12).
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-VAL-DESC     PIC X(30).
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-VAL-QTY      PIC ZZZ,ZZZ,ZZ9.999.
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-VAL-COST     PIC ZZZ,ZZZ,ZZ9.9999.
               10  FILLER          PIC X(2) VALUE SPACES.
               10  RL-VAL-VALUE    PIC ZZZ,ZZZ,ZZZ,ZZ9.99.
               10  FILLER          PIC X(15) VALUE SPACES.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-TRANSACTIONS
               UNTIL TXN-EOF
           PERFORM 3000-GENERATE-REORDER-REPORT
           PERFORM 4000-GENERATE-VALUATION-REPORT
           PERFORM 5000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           OPEN I-O    INVENTORY-MASTER
           OPEN INPUT  TRANSACTION-FILE
           OPEN OUTPUT REORDER-REPORT
           OPEN OUTPUT VALUATION-REPORT
           PERFORM 1100-LOAD-CATEGORIES
           PERFORM 1200-READ-TRANSACTION.

       1100-LOAD-CATEGORIES.
           MOVE 'ELEC'   TO CAT-CODE(1)
           MOVE 'ELECTRONICS'          TO CAT-DESC(1)
           MOVE 'MECH'   TO CAT-CODE(2)
           MOVE 'MECHANICAL PARTS'     TO CAT-DESC(2)
           MOVE 'CHEM'   TO CAT-CODE(3)
           MOVE 'CHEMICALS'            TO CAT-DESC(3)
           MOVE 'PACK'   TO CAT-CODE(4)
           MOVE 'PACKAGING MATERIALS'  TO CAT-DESC(4)
           MOVE 'TOOL'   TO CAT-CODE(5)
           MOVE 'TOOLS & EQUIPMENT'    TO CAT-DESC(5)
           MOVE 'CONS'   TO CAT-CODE(6)
           MOVE 'CONSUMABLES'          TO CAT-DESC(6).

       1200-READ-TRANSACTION.
           READ TRANSACTION-FILE
               AT END MOVE '10' TO WS-TXN-STATUS
           END-READ.

       2000-PROCESS-TRANSACTIONS.
           PERFORM 2100-HASH-ITEM-CODE
           MOVE WS-ITEM-HASH TO WS-REL-KEY
           READ INVENTORY-MASTER
               INVALID KEY
                   MOVE 'Y' TO WS-ERROR-FLAG
                   ADD 1 TO WS-ERRORS
           END-READ
           IF NO-ERROR
               EVALUATE TRUE
                   WHEN TXN-RECEIPT
                       PERFORM 2200-PROCESS-RECEIPT
                   WHEN TXN-ISSUE
                       PERFORM 2300-PROCESS-ISSUE
                   WHEN TXN-ADJUST-UP OR TXN-ADJUST-DOWN
                       PERFORM 2400-PROCESS-ADJUSTMENT
                   WHEN TXN-RETURN
                       PERFORM 2500-PROCESS-RETURN
                   WHEN TXN-SCRAP
                       PERFORM 2600-PROCESS-SCRAP
               END-EVALUATE
               REWRITE INVENTORY-RECORD
                   INVALID KEY ADD 1 TO WS-ERRORS
               END-REWRITE
           END-IF
           MOVE 'N' TO WS-ERROR-FLAG
           PERFORM 1200-READ-TRANSACTION.

       2100-HASH-ITEM-CODE.
           *> Simple hash: sum of character values mod table size
           MOVE ZEROS TO WS-ITEM-HASH
           INSPECT TXN-ITEM-CODE
               TALLYING WS-ITEM-HASH FOR ALL CHARACTERS
           COMPUTE WS-ITEM-HASH =
               FUNCTION MOD(WS-ITEM-HASH, 99991) + 1.

       2200-PROCESS-RECEIPT.
           ADD TXN-QUANTITY TO INV-QTY-ON-HAND
           ADD TXN-QUANTITY TO INV-YTD-RECEIPTS
           SUBTRACT TXN-QUANTITY FROM INV-QTY-ON-ORDER
           MOVE TXN-DATE TO INV-LAST-RECEIPT
           ADD 1 TO WS-RECEIPTS-COUNT

           *> Add FIFO layer
           IF INV-FIFO-LAYERS < 20
               ADD 1 TO INV-FIFO-LAYERS
               MOVE TXN-DATE     TO FIFO-RECEIPT-DATE(INV-FIFO-LAYERS)
               MOVE TXN-QUANTITY TO FIFO-QTY(INV-FIFO-LAYERS)
               MOVE TXN-UNIT-COST TO FIFO-UNIT-COST(INV-FIFO-LAYERS)
           END-IF

           *> Recalculate average cost
           PERFORM 2210-CALC-AVG-COST.

       2210-CALC-AVG-COST.
           MOVE ZEROS TO WS-FIFO-VALUE
           PERFORM VARYING WS-LAYER-IDX FROM 1 BY 1
               UNTIL WS-LAYER-IDX > INV-FIFO-LAYERS
               COMPUTE WS-FIFO-VALUE = WS-FIFO-VALUE +
                   FIFO-QTY(WS-LAYER-IDX) *
                   FIFO-UNIT-COST(WS-LAYER-IDX)
           END-PERFORM
           IF INV-QTY-ON-HAND > ZEROS
               COMPUTE INV-AVG-COST ROUNDED =
                   WS-FIFO-VALUE / INV-QTY-ON-HAND
           END-IF
           COMPUTE INV-TOTAL-VALUE = INV-QTY-ON-HAND * INV-AVG-COST.

       2300-PROCESS-ISSUE.
           IF TXN-QUANTITY > INV-QTY-ON-HAND
               ADD 1 TO WS-ERRORS
               MOVE 'Y' TO WS-ERROR-FLAG
           ELSE
               SUBTRACT TXN-QUANTITY FROM INV-QTY-ON-HAND
               ADD TXN-QUANTITY TO INV-YTD-ISSUES
               MOVE TXN-DATE TO INV-LAST-ISSUE
               ADD 1 TO WS-ISSUES-COUNT

               *> FIFO cost relief
               MOVE TXN-QUANTITY TO WS-ISSUE-QTY-REMAIN
               MOVE ZEROS TO WS-ISSUE-COST
               MOVE 1 TO WS-LAYER-IDX
               PERFORM UNTIL WS-ISSUE-QTY-REMAIN <= ZEROS
                   OR WS-LAYER-IDX > INV-FIFO-LAYERS
                   IF FIFO-QTY(WS-LAYER-IDX) <= WS-ISSUE-QTY-REMAIN
                       COMPUTE WS-ISSUE-COST = WS-ISSUE-COST +
                           FIFO-QTY(WS-LAYER-IDX) *
                           FIFO-UNIT-COST(WS-LAYER-IDX)
                       SUBTRACT FIFO-QTY(WS-LAYER-IDX)
                           FROM WS-ISSUE-QTY-REMAIN
                       MOVE ZEROS TO FIFO-QTY(WS-LAYER-IDX)
                   ELSE
                       COMPUTE WS-ISSUE-COST = WS-ISSUE-COST +
                           WS-ISSUE-QTY-REMAIN *
                           FIFO-UNIT-COST(WS-LAYER-IDX)
                       SUBTRACT WS-ISSUE-QTY-REMAIN
                           FROM FIFO-QTY(WS-LAYER-IDX)
                       MOVE ZEROS TO WS-ISSUE-QTY-REMAIN
                   END-IF
                   ADD 1 TO WS-LAYER-IDX
               END-PERFORM
               PERFORM 2210-CALC-AVG-COST
           END-IF.

       2400-PROCESS-ADJUSTMENT.
           ADD TXN-QUANTITY TO INV-QTY-ON-HAND
           ADD 1 TO WS-ADJUSTMENTS
           PERFORM 2210-CALC-AVG-COST.

       2500-PROCESS-RETURN.
           ADD TXN-QUANTITY TO INV-QTY-ON-HAND
           SUBTRACT TXN-QUANTITY FROM INV-YTD-ISSUES
           PERFORM 2210-CALC-AVG-COST.

       2600-PROCESS-SCRAP.
           SUBTRACT TXN-QUANTITY FROM INV-QTY-ON-HAND
           ADD 1 TO WS-ADJUSTMENTS
           PERFORM 2210-CALC-AVG-COST.

       3000-GENERATE-REORDER-REPORT.
           WRITE REORDER-LINE FROM
               "ITEM CODE    DESCRIPTION                    SUPPLIER" &
               "   ON HAND        REORDER PT   SUGGEST QTY  LEAD"
           WRITE REORDER-LINE FROM ALL '-'

           MOVE LOW-VALUES TO WS-REL-KEY
           PERFORM 3100-REORDER-SCAN
               UNTIL INV-EOF

           MOVE SPACES TO REORDER-LINE
           STRING 'Items requiring reorder: ' WS-REORDER-COUNT
               DELIMITED SIZE INTO REORDER-LINE
           WRITE REORDER-LINE.

       3100-REORDER-SCAN.
           READ INVENTORY-MASTER NEXT
               AT END MOVE '10' TO WS-INV-STATUS
           END-READ
           IF NOT INV-EOF
               COMPUTE WS-FIFO-VALUE =
                   INV-QTY-ON-HAND + INV-QTY-ON-ORDER
               IF WS-FIFO-VALUE <= INV-REORDER-POINT
                   ADD 1 TO WS-REORDER-COUNT
                   MOVE INV-ITEM-CODE    TO RL-RO-ITEM
                   MOVE INV-DESCRIPTION  TO RL-RO-DESC
                   MOVE INV-SUPPLIER-CODE TO RL-RO-SUPPLIER
                   MOVE INV-QTY-ON-HAND  TO RL-RO-ON-HAND
                   MOVE INV-REORDER-POINT TO RL-RO-REORDER
                   MOVE INV-REORDER-QTY  TO RL-RO-SUGGEST
                   MOVE INV-LEAD-TIME-DAYS TO RL-RO-LEAD
                   WRITE REORDER-LINE FROM RL-REORDER
               END-IF
           END-IF.

       4000-GENERATE-VALUATION-REPORT.
           WRITE VALUATION-LINE FROM
               "ITEM CODE    DESCRIPTION                    QTY ON HAND" &
               "     AVG COST       TOTAL VALUE"
           WRITE VALUATION-LINE FROM ALL '-'

           MOVE LOW-VALUES TO WS-REL-KEY
           MOVE '00' TO WS-INV-STATUS
           PERFORM 4100-VALUATION-SCAN
               UNTIL INV-EOF

           WRITE VALUATION-LINE FROM ALL '='
           MOVE SPACES TO VALUATION-LINE
           STRING 'TOTAL INVENTORY VALUE: ' WS-TOTAL-INV-VALUE
               DELIMITED SIZE INTO VALUATION-LINE
           WRITE VALUATION-LINE

           *> Category breakdown
           WRITE VALUATION-LINE FROM SPACES
           WRITE VALUATION-LINE FROM 'VALUATION BY CATEGORY:'
           PERFORM VARYING CAT-IDX FROM 1 BY 1
               UNTIL CAT-IDX > 20
               IF CAT-ITEM-COUNT(CAT-IDX) > ZEROS
                   MOVE SPACES TO VALUATION-LINE
                   STRING CAT-DESC(CAT-IDX) ': '
                          CAT-ITEM-COUNT(CAT-IDX) ' items, value: '
                          CAT-VALUE(CAT-IDX)
                       DELIMITED SIZE INTO VALUATION-LINE
                   WRITE VALUATION-LINE
               END-IF
           END-PERFORM.

       4100-VALUATION-SCAN.
           READ INVENTORY-MASTER NEXT
               AT END MOVE '10' TO WS-INV-STATUS
           END-READ
           IF NOT INV-EOF
               MOVE INV-ITEM-CODE    TO RL-VAL-ITEM
               MOVE INV-DESCRIPTION  TO RL-VAL-DESC
               MOVE INV-QTY-ON-HAND  TO RL-VAL-QTY
               MOVE INV-AVG-COST     TO RL-VAL-COST
               MOVE INV-TOTAL-VALUE  TO RL-VAL-VALUE
               WRITE VALUATION-LINE FROM RL-VALUATION
               ADD INV-TOTAL-VALUE TO WS-TOTAL-INV-VALUE

               *> Accumulate by category
               PERFORM VARYING CAT-IDX FROM 1 BY 1
                   UNTIL CAT-IDX > 20
                   IF CAT-CODE(CAT-IDX) = INV-CATEGORY
                       ADD INV-TOTAL-VALUE TO CAT-VALUE(CAT-IDX)
                       ADD 1 TO CAT-ITEM-COUNT(CAT-IDX)
                   END-IF
               END-PERFORM
           END-IF.

       5000-PRINT-SUMMARY.
           DISPLAY "=== INVENTORY PROCESSING COMPLETE ==="
           DISPLAY "Receipts processed  : " WS-RECEIPTS-COUNT
           DISPLAY "Issues processed    : " WS-ISSUES-COUNT
           DISPLAY "Adjustments         : " WS-ADJUSTMENTS
           DISPLAY "Errors              : " WS-ERRORS
           DISPLAY "Items to reorder    : " WS-REORDER-COUNT
           DISPLAY "Total inventory value: " WS-TOTAL-INV-VALUE.

       9000-TERMINATE.
           CLOSE INVENTORY-MASTER
                 TRANSACTION-FILE
                 REORDER-REPORT
                 VALUATION-REPORT.
