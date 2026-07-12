use super::helpers::compile_ok;

// ── SEARCH ALL — binary search on sorted table ────────────────

#[test]
fn search_all_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-code
               INDEXED BY ws-idx.
               10 ws-code  PIC 9(3).
               10 ws-label PIC X(10).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 100 TO ws-code(1)  MOVE "Alpha"    TO ws-label(1)
           MOVE 200 TO ws-code(2)  MOVE "Beta"     TO ws-label(2)
           MOVE 300 TO ws-code(3)  MOVE "Gamma"    TO ws-label(3)
           MOVE 400 TO ws-code(4)  MOVE "Delta"    TO ws-label(4)
           MOVE 500 TO ws-code(5)  MOVE "Epsilon"  TO ws-label(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-code(ws-idx) = 300
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_not_found() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-id
               INDEXED BY ws-idx.
               10 ws-id  PIC 9(4).
               10 ws-val PIC X(5).
       01 ws-result PIC X(10) VALUE "not found".
       PROCEDURE DIVISION.
           MOVE 1010 TO ws-id(1)
           MOVE 2020 TO ws-id(2)
           MOVE 3030 TO ws-id(3)
           MOVE 4040 TO ws-id(4)
           MOVE 5050 TO ws-id(5)
           SEARCH ALL ws-entry
               AT END MOVE "missing"   TO ws-result
               WHEN ws-id(ws-idx) = 9999
                   MOVE "found"        TO ws-result
           END-SEARCH
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_first_element() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 10 TIMES
               ASCENDING KEY IS ws-key
               INDEXED BY ws-idx.
               10 ws-key  PIC 9(2).
               10 ws-data PIC X(5).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 10 TO ws-key(1)   MOVE 20 TO ws-key(2)
           MOVE 30 TO ws-key(3)   MOVE 40 TO ws-key(4)
           MOVE 50 TO ws-key(5)   MOVE 60 TO ws-key(6)
           MOVE 70 TO ws-key(7)   MOVE 80 TO ws-key(8)
           MOVE 90 TO ws-key(9)   MOVE 99 TO ws-key(10)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-key(ws-idx) = 10
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_last_element() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 10 TIMES
               ASCENDING KEY IS ws-key
               INDEXED BY ws-idx.
               10 ws-key  PIC 9(2).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 10 TO ws-key(1)   MOVE 20 TO ws-key(2)
           MOVE 30 TO ws-key(3)   MOVE 40 TO ws-key(4)
           MOVE 50 TO ws-key(5)   MOVE 60 TO ws-key(6)
           MOVE 70 TO ws-key(7)   MOVE 80 TO ws-key(8)
           MOVE 90 TO ws-key(9)   MOVE 99 TO ws-key(10)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-key(ws-idx) = 99
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_alpha_key() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-name
               INDEXED BY ws-idx.
               10 ws-name  PIC X(10).
               10 ws-score PIC 99.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MOVE "Alice"    TO ws-name(1)  MOVE 85 TO ws-score(1)
           MOVE "Bob"      TO ws-name(2)  MOVE 92 TO ws-score(2)
           MOVE "Charlie"  TO ws-name(3)  MOVE 78 TO ws-score(3)
           MOVE "Diana"    TO ws-name(4)  MOVE 95 TO ws-score(4)
           MOVE "Eve"      TO ws-name(5)  MOVE 88 TO ws-score(5)
           SEARCH ALL ws-entry
               AT END MOVE 0 TO ws-result
               WHEN ws-name(ws-idx) = "Charlie"
                   MOVE ws-score(ws-idx) TO ws-result
           END-SEARCH
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_compound_key() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 6 TIMES
               ASCENDING KEY IS ws-dept ws-emp-id
               INDEXED BY ws-idx.
               10 ws-dept   PIC X(3).
               10 ws-emp-id PIC 9(4).
               10 ws-salary PIC 9(6).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "ACC" TO ws-dept(1)  MOVE 1001 TO ws-emp-id(1)
           MOVE "ACC" TO ws-dept(2)  MOVE 1002 TO ws-emp-id(2)
           MOVE "ENG" TO ws-dept(3)  MOVE 2001 TO ws-emp-id(3)
           MOVE "ENG" TO ws-dept(4)  MOVE 2002 TO ws-emp-id(4)
           MOVE "MKT" TO ws-dept(5)  MOVE 3001 TO ws-emp-id(5)
           MOVE "MKT" TO ws-dept(6)  MOVE 3002 TO ws-emp-id(6)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-dept(ws-idx) = "ENG" AND
                    ws-emp-id(ws-idx) = 2002
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_with_perform_action() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-product OCCURS 8 TIMES
               ASCENDING KEY IS ws-prod-id
               INDEXED BY ws-idx.
               10 ws-prod-id   PIC 9(5).
               10 ws-prod-name PIC X(20).
               10 ws-price     PIC 9(5)V99.
       01 ws-found-price PIC 9(5)V99 VALUE 0.
       01 ws-found-flag  PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 10001 TO ws-prod-id(1) MOVE  5.99 TO ws-price(1)
           MOVE 10002 TO ws-prod-id(2) MOVE 12.50 TO ws-price(2)
           MOVE 10003 TO ws-prod-id(3) MOVE  3.25 TO ws-price(3)
           MOVE 10004 TO ws-prod-id(4) MOVE 99.00 TO ws-price(4)
           MOVE 10005 TO ws-prod-id(5) MOVE 14.75 TO ws-price(5)
           MOVE 10006 TO ws-prod-id(6) MOVE  7.49 TO ws-price(6)
           MOVE 10007 TO ws-prod-id(7) MOVE 22.00 TO ws-price(7)
           MOVE 10008 TO ws-prod-id(8) MOVE  1.99 TO ws-price(8)
           SEARCH ALL ws-product
               AT END
                   MOVE "N" TO ws-found-flag
               WHEN ws-prod-id(ws-idx) = 10005
                   MOVE ws-price(ws-idx) TO ws-found-price
                   MOVE "Y" TO ws-found-flag
           END-SEARCH
           IF ws-found-flag = "Y"
               DISPLAY ws-found-price
           ELSE
               DISPLAY "not found"
           END-IF
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_descending_key() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               DESCENDING KEY IS ws-priority
               INDEXED BY ws-idx.
               10 ws-priority PIC 9.
               10 ws-task     PIC X(15).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 9 TO ws-priority(1)  MOVE "Critical" TO ws-task(1)
           MOVE 7 TO ws-priority(2)  MOVE "High"     TO ws-task(2)
           MOVE 5 TO ws-priority(3)  MOVE "Medium"   TO ws-task(3)
           MOVE 3 TO ws-priority(4)  MOVE "Low"      TO ws-task(4)
           MOVE 1 TO ws-priority(5)  MOVE "Deferred" TO ws-task(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-priority(ws-idx) = 5
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_varying_occurs() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 99 VALUE 5.
       01 ws-table.
           05 ws-entry OCCURS 1 TO 20 TIMES
               DEPENDING ON ws-count
               ASCENDING KEY IS ws-val
               INDEXED BY ws-idx.
               10 ws-val PIC 9(3).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE 5 TO ws-count
           MOVE 10 TO ws-val(1)
           MOVE 20 TO ws-val(2)
           MOVE 30 TO ws-val(3)
           MOVE 40 TO ws-val(4)
           MOVE 50 TO ws-val(5)
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-val(ws-idx) = 30
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_nested_table() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-outer.
           05 ws-region OCCURS 3 TIMES.
               10 ws-region-code PIC X(3).
               10 ws-city-table.
                   15 ws-city OCCURS 4 TIMES
                       ASCENDING KEY IS ws-city-id
                       INDEXED BY ws-ci.
                       20 ws-city-id   PIC 9(3).
                       20 ws-city-name PIC X(15).
       01 ws-found PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "NE" TO ws-region-code(1)
           MOVE 101 TO ws-city-id(1, 1)
           MOVE 102 TO ws-city-id(1, 2)
           MOVE 103 TO ws-city-id(1, 3)
           MOVE 104 TO ws-city-id(1, 4)
           SEARCH ALL ws-city(1)
               AT END MOVE "N" TO ws-found
               WHEN ws-city-id(1, ws-ci) = 103
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_index_after_search() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 5 TIMES
               ASCENDING KEY IS ws-num
               INDEXED BY ws-idx.
               10 ws-num PIC 9(3).
       01 ws-idx-val PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 100 TO ws-num(1)
           MOVE 200 TO ws-num(2)
           MOVE 300 TO ws-num(3)
           MOVE 400 TO ws-num(4)
           MOVE 500 TO ws-num(5)
           SEARCH ALL ws-entry
               AT END MOVE 0 TO ws-idx-val
               WHEN ws-num(ws-idx) = 300
                   SET ws-idx-val TO ws-idx
           END-SEARCH
           DISPLAY ws-idx-val
           STOP RUN.
"#,
    );
}

#[test]
fn search_all_large_table() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-table.
           05 ws-entry OCCURS 100 TIMES
               ASCENDING KEY IS ws-code
               INDEXED BY ws-idx.
               10 ws-code  PIC 9(5).
               10 ws-value PIC X(10).
       01 ws-found PIC X VALUE "N".
       01 ws-i PIC 9(3).
       PROCEDURE DIVISION.
           PERFORM VARYING ws-i FROM 1 BY 1 UNTIL ws-i > 100
               MULTIPLY ws-i BY 10 GIVING ws-code(ws-i)
           END-PERFORM
           SEARCH ALL ws-entry
               AT END MOVE "N" TO ws-found
               WHEN ws-code(ws-idx) = 500
                   MOVE "Y" TO ws-found
           END-SEARCH
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}
