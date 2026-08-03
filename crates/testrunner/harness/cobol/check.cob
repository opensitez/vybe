      * Vybe test harness — COBOL.
      *
      * COBOL has no callable assertion helper worth splicing: there are no
      * parameterised paragraphs, so a `PERFORM` cannot take the value to
      * check. The assertion is therefore emitted INLINE after each DISPLAY,
      * and this file documents the shape the emitter produces rather than
      * providing code to splice.
      *
      * The emitted check, after `DISPLAY WS-A.`:
      *
      *     IF WS-A NOT = 6
      *         DISPLAY "FAIL: want [006] got [" WS-A "]"
      *         MOVE 1 TO RETURN-CODE
      *         RAISE EXCEPTION EC-PROGRAM
      *     END-IF.
      *
      * Two measured constraints shape it, both verified against `cobc`:
      *
      * 1. FAILURE MUST BE SIGNALLED TWICE, once per runtime.
      *    `STOP RUN WITH ERROR STATUS` is NOT in Vybe's COBOL grammar, so it
      *    is a PARSE error — the first version of this harness "failed" every
      *    test for that reason rather than by assertion. `MOVE 1 TO
      *    RETURN-CODE` exits 1 under cobc but 0 under Vybe (the VM has no
      *    exit-code path at all). `RAISE EXCEPTION EC-PROGRAM` throws under
      *    Vybe but is an unimplemented no-op under cobc. Emitting BOTH gives
      *    a correct non-zero verdict in each.
      *
      * 2. THE COMPARISON MUST MATCH THE OPERAND'S CLASS.
      *    Under Vybe, `WS-A = "006"` is FALSE for a `PIC 9(3)` holding 6,
      *    and moving it to a `PIC X` first does not help. Only
      *    numeric-vs-numeric (`WS-A = 6`) and alphanumeric-vs-alphanumeric
      *    (`WS-S = "hi"`) agree with cobc. The emitter therefore writes an
      *    unquoted numeric literal when the expected text is all digits and
      *    a quoted literal otherwise.
      *
      * The FAIL line is DISPLAYed before stopping because an uncaught error
      * renders as `RuntimeError: [object]` under Vybe, so the expected and
      * actual values would otherwise be lost.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. VYBECHK.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-A PIC 9(3) VALUE 6.
       PROCEDURE DIVISION.
           DISPLAY WS-A.
           IF WS-A NOT = 6
               DISPLAY "FAIL: want [006] got [" WS-A "]"
               MOVE 1 TO RETURN-CODE
               RAISE EXCEPTION EC-PROGRAM
           END-IF.
           STOP RUN.
