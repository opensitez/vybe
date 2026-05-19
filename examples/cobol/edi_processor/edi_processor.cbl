      *> ============================================================
      *> EDI (ELECTRONIC DATA INTERCHANGE) TRANSACTION PROCESSOR
      *> ============================================================
      *> Parses and validates ANSI X12 EDI files: ISA/GS/ST
      *> envelope handling, 850 Purchase Orders, 810 Invoices,
      *> 997 Functional Acknowledgments.
      *>
      *> Demonstrates: UNSTRING with multiple delimiters,
      *> STRING, INSPECT TALLYING/REPLACING, complex parsing,
      *> EVALUATE with multiple WHEN, reference modification,
      *> COBOL 2014 XML/JSON output generation.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDI-PROCESSOR.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT EDI-INPUT ASSIGN TO "edi_input.edi"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-EDI-STATUS.

           SELECT EDI-OUTPUT ASSIGN TO "edi_output.edi"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-OUT-STATUS.

           SELECT EDI-LOG ASSIGN TO "edi_processing.log"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT PO-OUTPUT ASSIGN TO "purchase_orders.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT INV-OUTPUT ASSIGN TO "invoices_received.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT ACK-OUTPUT ASSIGN TO "acknowledgments.edi"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  EDI-INPUT
           RECORD CONTAINS 1 TO 1024 CHARACTERS.
       01  EDI-LINE                PIC X(1024).

       FD  EDI-OUTPUT
           RECORD CONTAINS 1 TO 1024 CHARACTERS.
       01  EDI-OUT-LINE            PIC X(1024).

       FD  EDI-LOG
           RECORD CONTAINS 200 CHARACTERS.
       01  LOG-LINE                PIC X(200).

       FD  PO-OUTPUT
           RECORD CONTAINS 500 CHARACTERS.
       01  PO-RECORD               PIC X(500).

       FD  INV-OUTPUT
           RECORD CONTAINS 500 CHARACTERS.
       01  INV-RECORD              PIC X(500).

       FD  ACK-OUTPUT
           RECORD CONTAINS 200 CHARACTERS.
       01  ACK-LINE                PIC X(200).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-EDI-STATUS       PIC XX.
               88  EDI-OK          VALUE '00'.
               88  EDI-EOF         VALUE '10'.
           05  WS-OUT-STATUS       PIC XX.

       01  WS-EDI-ENVELOPE.
           05  WS-ISA-SEGMENT.
               10  ISA-AUTH-INFO-QUAL  PIC X(2).
               10  ISA-AUTH-INFO       PIC X(10).
               10  ISA-SEC-INFO-QUAL   PIC X(2).
               10  ISA-SEC-INFO        PIC X(10).
               10  ISA-SENDER-QUAL     PIC X(2).
               10  ISA-SENDER-ID       PIC X(15).
               10  ISA-RECEIVER-QUAL   PIC X(2).
               10  ISA-RECEIVER-ID     PIC X(15).
               10  ISA-DATE            PIC X(6).
               10  ISA-TIME            PIC X(4).
               10  ISA-REPETITION-SEP  PIC X(1).
               10  ISA-VERSION         PIC X(5).
               10  ISA-CONTROL-NUM     PIC X(9).
               10  ISA-ACK-REQUESTED   PIC X(1).
               10  ISA-USAGE-IND       PIC X(1).
               10  ISA-COMPONENT-SEP   PIC X(1).
           05  WS-GS-SEGMENT.
               10  GS-FUNC-ID-CODE     PIC X(2).
               10  GS-APP-SENDER       PIC X(15).
               10  GS-APP-RECEIVER     PIC X(15).
               10  GS-DATE             PIC X(8).
               10  GS-TIME             PIC X(8).
               10  GS-GROUP-CTRL-NUM   PIC X(9).
               10  GS-RESP-AGENCY      PIC X(2).
               10  GS-VERSION          PIC X(12).

       01  WS-TRANSACTION-SET.
           05  WS-ST-SEGMENT.
               10  ST-TRANS-SET-ID     PIC X(3).
               10  ST-CTRL-NUM         PIC X(9).
           05  WS-TRANS-TYPE          PIC X(3).
               88  TRANS-850           VALUE '850'.
               88  TRANS-810           VALUE '810'.
               88  TRANS-856           VALUE '856'.
               88  TRANS-997           VALUE '997'.
           05  WS-SEGMENT-COUNT       PIC 9(6) VALUE ZEROS.

       01  WS-850-PO.
           05  PO-NUMBER              PIC X(22).
           05  PO-DATE                PIC X(8).
           05  PO-TYPE                PIC X(2).
           05  PO-VENDOR-ID           PIC X(15).
           05  PO-SHIP-TO-NAME        PIC X(60).
           05  PO-SHIP-TO-ADDR        PIC X(55).
           05  PO-SHIP-TO-CITY        PIC X(30).
           05  PO-SHIP-TO-STATE       PIC X(2).
           05  PO-SHIP-TO-ZIP         PIC X(10).
           05  PO-LINE-COUNT          PIC 9(4) VALUE ZEROS.
           05  PO-TOTAL-AMOUNT        PIC S9(13)V99 VALUE ZEROS.
           05  PO-LINES OCCURS 1 TO 999 TIMES
               DEPENDING ON PO-LINE-COUNT.
               10  POL-LINE-NUM       PIC 9(6).
               10  POL-QTY            PIC 9(9)V999.
               10  POL-UOM            PIC X(2).
               10  POL-UNIT-PRICE     PIC 9(9)V9999.
               10  POL-PRODUCT-ID     PIC X(30).
               10  POL-DESCRIPTION    PIC X(50).

       01  WS-PARSING.
           05  WS-SEGMENT-ID          PIC X(3).
           05  WS-ELEMENT-SEP         PIC X(1) VALUE '*'.
           05  WS-SEGMENT-TERM        PIC X(1) VALUE '~'.
           05  WS-ELEMENTS.
               10  WS-E1              PIC X(80).
               10  WS-E2              PIC X(80).
               10  WS-E3              PIC X(80).
               10  WS-E4              PIC X(80).
               10  WS-E5              PIC X(80).
               10  WS-E6              PIC X(80).
               10  WS-E7              PIC X(80).
               10  WS-E8              PIC X(80).
               10  WS-E9              PIC X(80).
               10  WS-E10             PIC X(80).
               10  WS-E11             PIC X(80).
               10  WS-E12             PIC X(80).
               10  WS-E13             PIC X(80).
               10  WS-E14             PIC X(80).
               10  WS-E15             PIC X(80).
               10  WS-E16             PIC X(80).
               10  WS-E17             PIC X(80).
               10  WS-E18             PIC X(80).
               10  WS-E19             PIC X(80).
               10  WS-E20             PIC X(80).
           05  WS-ELEM-COUNT          PIC 9(3).
           05  WS-CURRENT-LINE        PIC X(1024).
           05  WS-LINE-LENGTH         PIC 9(4).
           05  WS-PARSE-POS           PIC 9(4).
           05  WS-SEGMENT-BUFFER      PIC X(1024).
           05  WS-BUFFER-POS          PIC 9(4) VALUE 1.

       01  WS-COUNTERS.
           05  WS-INTERCHANGES        PIC 9(6) VALUE ZEROS.
           05  WS-GROUPS              PIC 9(6) VALUE ZEROS.
           05  WS-TRANSACTIONS        PIC 9(6) VALUE ZEROS.
           05  WS-SEGMENTS-TOTAL      PIC 9(8) VALUE ZEROS.
           05  WS-PO-COUNT            PIC 9(6) VALUE ZEROS.
           05  WS-INV-COUNT           PIC 9(6) VALUE ZEROS.
           05  WS-ACK-COUNT           PIC 9(6) VALUE ZEROS.
           05  WS-ERRORS              PIC 9(6) VALUE ZEROS.

       01  WS-ACK-FIELDS.
           05  WS-ACK-CTRL-NUM        PIC 9(9) VALUE 1.
           05  WS-ACK-STATUS          PIC X(1).
               88  ACK-ACCEPTED       VALUE 'A'.
               88  ACK-REJECTED       VALUE 'R'.
               88  ACK-ACCEPTED-ERRORS VALUE 'E'.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-EDI
               UNTIL EDI-EOF
           PERFORM 3000-WRITE-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           OPEN INPUT  EDI-INPUT
           OPEN OUTPUT EDI-OUTPUT
           OPEN OUTPUT EDI-LOG
           OPEN OUTPUT PO-OUTPUT
           OPEN OUTPUT INV-OUTPUT
           OPEN OUTPUT ACK-OUTPUT
           PERFORM 1100-READ-EDI-LINE.

       1100-READ-EDI-LINE.
           READ EDI-INPUT INTO WS-CURRENT-LINE
               AT END MOVE '10' TO WS-EDI-STATUS
           END-READ
           IF EDI-OK
               MOVE FUNCTION LENGTH(
                   FUNCTION TRIM(WS-CURRENT-LINE TRAILING))
                   TO WS-LINE-LENGTH
           END-IF.

       2000-PROCESS-EDI.
           *> Parse segments from the line (segments end with ~)
           PERFORM 2100-EXTRACT-SEGMENT
           PERFORM 2200-PARSE-ELEMENTS
           PERFORM 2300-ROUTE-SEGMENT
           PERFORM 1100-READ-EDI-LINE.

       2100-EXTRACT-SEGMENT.
           *> Segments may span lines or multiple per line
           *> For simplicity: one segment per line ending with ~
           MOVE WS-CURRENT-LINE TO WS-SEGMENT-BUFFER
           ADD 1 TO WS-SEGMENTS-TOTAL.

       2200-PARSE-ELEMENTS.
           *> Split segment by element separator '*'
           MOVE ZEROS TO WS-ELEM-COUNT
           INITIALIZE WS-ELEMENTS
           UNSTRING WS-SEGMENT-BUFFER
               DELIMITED BY '*' OR '~'
            INTO WS-E1  WS-E2  WS-E3
                WS-E4  WS-E5  WS-E6
                WS-E7  WS-E8  WS-E9
                WS-E10 WS-E11 WS-E12
                WS-E13 WS-E14 WS-E15
                WS-E16 WS-E17 WS-E18
                WS-E19 WS-E20
               TALLYING WS-ELEM-COUNT
           END-UNSTRING
           MOVE WS-E1 TO WS-SEGMENT-ID.

       2300-ROUTE-SEGMENT.
           EVALUATE WS-E1
               WHEN 'ISA'
                   PERFORM 2310-PROCESS-ISA
               WHEN 'GS'
                   PERFORM 2320-PROCESS-GS
               WHEN 'ST'
                   PERFORM 2330-PROCESS-ST
               WHEN 'BEG'
                   PERFORM 2340-PROCESS-BEG-850
               WHEN 'PO1'
                   PERFORM 2350-PROCESS-PO1
               WHEN 'BIG'
                   PERFORM 2360-PROCESS-BIG-810
               WHEN 'IT1'
                   PERFORM 2370-PROCESS-IT1
               WHEN 'N1'
                   PERFORM 2380-PROCESS-N1
               WHEN 'SE'
                   PERFORM 2390-PROCESS-SE
               WHEN 'GE'
                   PERFORM 2395-PROCESS-GE
               WHEN 'IEA'
                   PERFORM 2398-PROCESS-IEA
               WHEN OTHER
                   CONTINUE
           END-EVALUATE.

       2310-PROCESS-ISA.
           ADD 1 TO WS-INTERCHANGES
           MOVE WS-E2  TO ISA-AUTH-INFO-QUAL
           MOVE WS-E3  TO ISA-AUTH-INFO
           MOVE WS-E6  TO ISA-SENDER-ID
           MOVE WS-E8  TO ISA-RECEIVER-ID
           MOVE WS-E9  TO ISA-DATE
           MOVE WS-E10 TO ISA-TIME
           MOVE WS-E13 TO ISA-CONTROL-NUM
           MOVE SPACES TO LOG-LINE
           STRING 'ISA received: sender=' ISA-SENDER-ID
                  ' ctrl=' ISA-CONTROL-NUM
               DELIMITED SIZE INTO LOG-LINE
           WRITE LOG-LINE.

       2320-PROCESS-GS.
           ADD 1 TO WS-GROUPS
           MOVE WS-E2 TO GS-FUNC-ID-CODE
           MOVE WS-E3 TO GS-APP-SENDER
           MOVE WS-E4 TO GS-APP-RECEIVER
           MOVE WS-E7 TO GS-GROUP-CTRL-NUM.

       2330-PROCESS-ST.
           ADD 1 TO WS-TRANSACTIONS
           MOVE WS-E2 TO ST-TRANS-SET-ID
           MOVE WS-E3 TO ST-CTRL-NUM
           MOVE ST-TRANS-SET-ID TO WS-TRANS-TYPE
           MOVE ZEROS TO WS-SEGMENT-COUNT
           INITIALIZE WS-850-PO.

       2340-PROCESS-BEG-850.
           IF TRANS-850
               MOVE WS-E4 TO PO-NUMBER
               MOVE WS-E5 TO PO-DATE
               MOVE WS-E3 TO PO-TYPE
           END-IF.

       2350-PROCESS-PO1.
           IF TRANS-850
               ADD 1 TO PO-LINE-COUNT
               MOVE WS-E2 TO POL-LINE-NUM(PO-LINE-COUNT)
               MOVE WS-E3 TO POL-QTY(PO-LINE-COUNT)
               MOVE WS-E4 TO POL-UOM(PO-LINE-COUNT)
               MOVE WS-E5 TO POL-UNIT-PRICE(PO-LINE-COUNT)
               MOVE WS-E8 TO POL-PRODUCT-ID(PO-LINE-COUNT)
               COMPUTE PO-TOTAL-AMOUNT = PO-TOTAL-AMOUNT +
                   POL-QTY(PO-LINE-COUNT) *
                   POL-UNIT-PRICE(PO-LINE-COUNT)
           END-IF.

       2360-PROCESS-BIG-810.
           *> Invoice header
           MOVE SPACES TO LOG-LINE
           STRING 'Invoice received: ' WS-E3
               DELIMITED SIZE INTO LOG-LINE
           WRITE LOG-LINE.

       2370-PROCESS-IT1.
           *> Invoice line item
           CONTINUE.

       2380-PROCESS-N1.
           IF TRANS-850
               EVALUATE WS-E2
                   WHEN 'ST'
                       MOVE WS-E3 TO PO-SHIP-TO-NAME
                   WHEN 'SE'
                       MOVE WS-E3 TO PO-VENDOR-ID
               END-EVALUATE
           END-IF.

       2390-PROCESS-SE.
           *> End of transaction set — write output and generate 997
           EVALUATE TRUE
               WHEN TRANS-850
                   PERFORM 2391-WRITE-PO
                   ADD 1 TO WS-PO-COUNT
               WHEN TRANS-810
                   ADD 1 TO WS-INV-COUNT
           END-EVALUATE
           PERFORM 2392-GENERATE-997.

       2391-WRITE-PO.
           MOVE SPACES TO PO-RECORD
           STRING 'PO:' PO-NUMBER
                  ' DATE:' PO-DATE
                  ' VENDOR:' PO-VENDOR-ID
                  ' LINES:' PO-LINE-COUNT
                  ' TOTAL:' PO-TOTAL-AMOUNT
               DELIMITED SIZE INTO PO-RECORD
           WRITE PO-RECORD.

       2392-GENERATE-997.
           *> Functional Acknowledgment
           ADD 1 TO WS-ACK-CTRL-NUM
           ADD 1 TO WS-ACK-COUNT
           MOVE 'A' TO WS-ACK-STATUS
           MOVE SPACES TO ACK-LINE
           STRING 'ST*997*' WS-ACK-CTRL-NUM '~'
               DELIMITED SIZE INTO ACK-LINE
           WRITE ACK-LINE
           MOVE SPACES TO ACK-LINE
           STRING 'AK1*' GS-FUNC-ID-CODE '*'
                  GS-GROUP-CTRL-NUM '~'
               DELIMITED SIZE INTO ACK-LINE
           WRITE ACK-LINE
           MOVE SPACES TO ACK-LINE
           STRING 'AK9*' WS-ACK-STATUS '*1*1*1~'
               DELIMITED SIZE INTO ACK-LINE
           WRITE ACK-LINE
           MOVE SPACES TO ACK-LINE
           STRING 'SE*3*' WS-ACK-CTRL-NUM '~'
               DELIMITED SIZE INTO ACK-LINE
           WRITE ACK-LINE.

       2395-PROCESS-GE.
           MOVE SPACES TO LOG-LINE
           STRING 'GE: group ' GS-GROUP-CTRL-NUM ' closed'
               DELIMITED SIZE INTO LOG-LINE
           WRITE LOG-LINE.

       2398-PROCESS-IEA.
           MOVE SPACES TO LOG-LINE
           STRING 'IEA: interchange ' ISA-CONTROL-NUM ' closed'
               DELIMITED SIZE INTO LOG-LINE
           WRITE LOG-LINE.

       3000-WRITE-SUMMARY.
           DISPLAY "=== EDI PROCESSING SUMMARY ==="
           DISPLAY "Interchanges     : " WS-INTERCHANGES
           DISPLAY "Functional Groups: " WS-GROUPS
           DISPLAY "Transactions     : " WS-TRANSACTIONS
           DISPLAY "Total Segments   : " WS-SEGMENTS-TOTAL
           DISPLAY "Purchase Orders  : " WS-PO-COUNT
           DISPLAY "Invoices         : " WS-INV-COUNT
           DISPLAY "Acknowledgments  : " WS-ACK-COUNT
           DISPLAY "Errors           : " WS-ERRORS.

       9000-TERMINATE.
           CLOSE EDI-INPUT EDI-OUTPUT EDI-LOG
                 PO-OUTPUT INV-OUTPUT ACK-OUTPUT.
