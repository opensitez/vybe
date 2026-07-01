use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn string_size_into_target_compiles() {
    compile_ok(&p(
        "01 A PIC X(3) VALUE \"ONE\".\n01 B PIC X(3) VALUE \"TWO\".\n01 R PIC X(10).",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R.",
    ));
}

#[test]
fn string_with_pointer_compiles() {
    compile_ok(&p(
        "01 A PIC X(2) VALUE \"AB\".\n01 B PIC X(2) VALUE \"CD\".\n01 R PIC X(10).\n01 P PIC 9(2) VALUE 1.",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R WITH POINTER P.",
    ));
}

#[test]
fn unstring_basic_comma_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"A,B,C\".\n01 F1 PIC X(3).\n01 F2 PIC X(3).\n01 F3 PIC X(3).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3.",
    ));
}

#[test]
fn unstring_with_count_and_delimiter_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"AA,BBB\".\n01 F1 PIC X(5).\n01 D1 PIC X.\n01 C1 PIC 9(2).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 DELIMITER IN D1 COUNT IN C1.",
    ));
}

#[test]
fn unstring_with_tallying_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"A,B,C\".\n01 F1 PIC X(2).\n01 F2 PIC X(2).\n01 F3 PIC X(2).\n01 T PIC 9 VALUE 0.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3 TALLYING IN T.",
    ));
}

#[test]
fn inspect_tallying_all_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(8) VALUE \"ABABXABA\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT TXT TALLYING C FOR ALL \"A\".",
    ));
}

#[test]
fn inspect_tallying_leading_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(8) VALUE \"00012345\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT TXT TALLYING C FOR LEADING \"0\".",
    ));
}

#[test]
fn inspect_replacing_first_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(6) VALUE \"AAAAAA\".",
        "    INSPECT TXT REPLACING FIRST \"A\" BY \"B\".",
    ));
}

#[test]
fn table_occurs_indexed_move_compiles() {
    compile_ok(&p(
        "01 TBL.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC 9(2).",
        "    SET I TO 1.\n    MOVE 12 TO K(I).",
    ));
}

#[test]
fn table_set_up_down_compiles() {
    compile_ok(&p(
        "01 TBL PIC 9 OCCURS 5 TIMES INDEXED BY I.",
        "    SET I TO 1.\n    SET I UP BY 2.\n    SET I DOWN BY 1.",
    ));
}

#[test]
fn search_linear_table_compiles() {
    compile_ok(&p(
        "01 TBL.\n   05 E OCCURS 3 TIMES INDEXED BY I.\n      10 K PIC 9.",
        "    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    SET I TO 1.\n    SEARCH E WHEN K(I) = 2 DISPLAY \"Y\" END-SEARCH.",
    ));
}

#[test]
fn search_all_table_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 TAB.\n   05 E OCCURS 4 TIMES ASCENDING KEY IS K INDEXED BY I.\n      10 K PIC 9(2).\nPROCEDURE DIVISION.\n    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    MOVE 4 TO K(4).\n    SEARCH ALL E WHEN K(I) = 3 DISPLAY \"Y\" END-SEARCH.\n    STOP RUN.",
    );
}