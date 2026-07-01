use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn json_generate_compiles() { compile_ok(&p("01 R.\n   05 N PIC X(10) VALUE \"A\".\n01 J PIC X(100).", "    JSON GENERATE J FROM R.")); }
#[test] fn json_parse_compiles() { compile_ok(&p("01 J PIC X(100).\n01 R PIC X(10).", "    JSON PARSE J INTO R.")); }
#[test] fn xml_generate_compiles() { compile_ok(&p("01 R PIC X(10) VALUE \"A\".\n01 X PIC X(200).\n01 L PIC 9(5).", "    XML GENERATE X FROM R COUNT IN L.")); }
#[test] fn xml_parse_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC X(100).\nPROCEDURE DIVISION.\n    XML PARSE X PROCESSING PROCEDURE H.\n    STOP RUN.\nH SECTION.\n    DISPLAY \"E\"."); }
#[test] fn sql_connect_compiles() { compile_ok(&p("01 D PIC X(100) VALUE \"sqlite:test.db\".", "    EXEC SQL CONNECT :D END-EXEC.")); }
#[test] fn sql_select_compiles() { compile_ok(&p("01 N PIC X(20).", "    EXEC SQL SELECT NAME INTO :N FROM USERS WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_insert_compiles() { compile_ok(&p("01 I PIC 9(5) VALUE 1.\n01 N PIC X(20) VALUE \"A\".", "    EXEC SQL INSERT INTO USERS (ID, NAME) VALUES (:I, :N) END-EXEC.")); }
#[test] fn sql_update_compiles() { compile_ok(&p("01 N PIC X(20) VALUE \"B\".", "    EXEC SQL UPDATE USERS SET NAME = :N WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_delete_compiles() { compile_ok(&p("", "    EXEC SQL DELETE FROM USERS WHERE ID = 1 END-EXEC.")); }
#[test] fn sql_commit_compiles() { compile_ok(&p("", "    EXEC SQL COMMIT END-EXEC.")); }
#[test] fn sql_rollback_compiles() { compile_ok(&p("", "    EXEC SQL ROLLBACK END-EXEC.")); }
#[test] fn sql_cursor_declare_compiles() { compile_ok(&p("", "    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.")); }
#[test] fn sql_cursor_open_compiles() { compile_ok(&p("", "    EXEC SQL OPEN C1 END-EXEC.")); }
#[test] fn sql_cursor_fetch_compiles() { compile_ok(&p("01 I PIC 9(5).", "    EXEC SQL FETCH C1 INTO :I END-EXEC.")); }
#[test] fn sql_cursor_close_compiles() { compile_ok(&p("", "    EXEC SQL CLOSE C1 END-EXEC.")); }
#[test] fn xml_generate_with_declaration_compiles() { compile_ok(&p("01 R PIC X(10) VALUE \"A\".\n01 X PIC X(200).\n01 L PIC 9(5).", "    XML GENERATE X FROM R COUNT IN L WITH XML-DECLARATION.")); }
#[test] fn json_parse_into_group_compiles() { compile_ok(&p("01 J PIC X(100).\n01 R.\n   05 N PIC X(10).", "    JSON PARSE J INTO R.")); }
