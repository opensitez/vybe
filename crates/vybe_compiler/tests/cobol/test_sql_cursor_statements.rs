use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn sql_connect_compiles() { compile_ok(&p("01 DSN PIC X(100) VALUE \"sqlite:test.db\".", "    EXEC SQL CONNECT :DSN END-EXEC.")); }
#[test] fn sql_disconnect_compiles() { compile_ok(&p("", "    EXEC SQL COMMIT END-EXEC.")); }
#[test] fn sql_select_into_compiles() { compile_ok(&p("01 N PIC X(50).", "    EXEC SQL SELECT NAME INTO :N FROM USERS WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_insert_values_compiles() { compile_ok(&p("01 I PIC 9(5) VALUE 1.\n01 N PIC X(20) VALUE \"A\".", "    EXEC SQL INSERT INTO USERS (ID, NAME) VALUES (:I, :N) END-EXEC.")); }
#[test] fn sql_update_values_compiles() { compile_ok(&p("01 N PIC X(20) VALUE \"B\".", "    EXEC SQL UPDATE USERS SET NAME = :N WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_delete_values_compiles() { compile_ok(&p("", "    EXEC SQL DELETE FROM USERS WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_commit_compiles() { compile_ok(&p("", "    EXEC SQL COMMIT END-EXEC.")); }
#[test] fn sql_rollback_compiles() { compile_ok(&p("", "    EXEC SQL ROLLBACK END-EXEC.")); }
#[test] fn sql_declare_cursor_compiles() { compile_ok(&p("", "    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.")); }
#[test] fn sql_open_cursor_compiles() { compile_ok(&p("", "    EXEC SQL OPEN C1 END-EXEC.")); }
#[test] fn sql_fetch_cursor_compiles() { compile_ok(&p("01 I PIC 9(5).", "    EXEC SQL FETCH C1 INTO :I END-EXEC.")); }
#[test] fn sql_close_cursor_compiles() { compile_ok(&p("", "    EXEC SQL CLOSE C1 END-EXEC.")); }
#[test] fn sql_cursor_loop_compiles() { compile_ok(&p("01 I PIC 9(5).\n01 SQLCODE PIC S9(9) VALUE 0.", "    EXEC SQL OPEN C1 END-EXEC.\n    PERFORM UNTIL SQLCODE NOT = 0\n        EXEC SQL FETCH C1 INTO :I END-EXEC\n    END-PERFORM.\n    EXEC SQL CLOSE C1 END-EXEC.")); }
#[test] fn sql_on_error_if_compiles() { compile_ok(&p("01 SQLCODE PIC S9(9) VALUE 0.", "    EXEC SQL SELECT 1 END-EXEC.\n    IF SQLCODE NOT = 0 DISPLAY \"E\" END-IF.")); }
#[test] fn sql_multi_statement_compiles() { compile_ok(&p("01 I PIC 9(5) VALUE 1.", "    EXEC SQL INSERT INTO T(ID) VALUES(:I) END-EXEC.\n    EXEC SQL UPDATE T SET ID = :I END-EXEC.")); }
#[test] fn sql_transaction_pattern_compiles() { compile_ok(&p("01 SQLCODE PIC S9(9) VALUE 0.", "    EXEC SQL INSERT INTO T(ID) VALUES(1) END-EXEC.\n    IF SQLCODE = 0 EXEC SQL COMMIT END-EXEC ELSE EXEC SQL ROLLBACK END-EXEC END-IF.")); }
#[test] fn sql_select_two_cols_compiles() { compile_ok(&p("01 A PIC X(20).\n01 B PIC 9(5).", "    EXEC SQL SELECT NAME, ID INTO :A, :B FROM USERS WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_cursor_two_cols_compiles() { compile_ok(&p("01 A PIC X(20).\n01 B PIC 9(5).", "    EXEC SQL FETCH C1 INTO :A, :B END-EXEC.")); }
