use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn add_two_operands_to_target() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 3.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A B TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn add_literal_to_target() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 1.",
        "    ADD 9 TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn add_giving_two_targets_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 7.\n01 B PIC 9(3) VALUE 8.\n01 R1 PIC 9(3).\n01 R2 PIC 9(3).",
        "    ADD A B GIVING R1 R2.",
    ));
}

#[test]
fn add_with_end_add_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 1.\n01 B PIC 9(3) VALUE 2.",
        "    ADD A TO B END-ADD.",
    ));
}

#[test]
fn subtract_literal_from_target() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 20.",
        "    SUBTRACT 6 FROM R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["14"]);
}

#[test]
fn subtract_giving_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 10.\n01 B PIC 9(3) VALUE 4.\n01 R PIC 9(3).",
        "    SUBTRACT B FROM A GIVING R.",
    ));
}

#[test]
fn subtract_with_end_subtract_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 9.\n01 B PIC 9(3) VALUE 2.",
        "    SUBTRACT B FROM A END-SUBTRACT.",
    ));
}

#[test]
fn multiply_literal_by_target() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 8.",
        "    MULTIPLY 3 BY R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["24"]);
}

#[test]
fn multiply_giving_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 6.\n01 B PIC 9(3) VALUE 7.\n01 R PIC 9(3).",
        "    MULTIPLY A BY B GIVING R.",
    ));
}

#[test]
fn multiply_with_end_multiply_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 2.\n01 B PIC 9(3) VALUE 5.",
        "    MULTIPLY A BY B END-MULTIPLY.",
    ));
}

#[test]
fn divide_literal_into_target_compiles() {
    compile_ok(&p("01 R PIC 9(3) VALUE 16.", "    DIVIDE 2 INTO R."));
}

#[test]
fn divide_by_giving_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 20.\n01 B PIC 9(3) VALUE 5.\n01 R PIC 9(3).",
        "    DIVIDE A BY B GIVING R.",
    ));
}

#[test]
fn divide_into_giving_remainder_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 7.\n01 B PIC 9(3) VALUE 3.\n01 Q PIC 9(3).\n01 M PIC 9(3).",
        "    DIVIDE B INTO A GIVING Q REMAINDER M.",
    ));
}

#[test]
fn divide_with_end_divide_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 12.\n01 B PIC 9(3) VALUE 3.",
        "    DIVIDE B INTO A END-DIVIDE.",
    ));
}

#[test]
fn compute_basic_sum() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = 1 + 2 + 3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["6"]);
}

#[test]
fn compute_nested_parentheses_compiles() {
    compile_ok(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = (2 + (3 * 4)) - 1.",
    ));
}

#[test]
fn if_compound_and_branch() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    IF A = 1 AND B = 2\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn if_compound_or_branch() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 2.",
        "    IF A = 1 OR B = 2\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn if_not_condition_branch() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.",
        "    IF NOT A = 1\n        DISPLAY \"N1\"\n    ELSE\n        DISPLAY \"N0\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["N1"]);
}

#[test]
fn evaluate_true_multiple_when() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 2.",
        "    EVALUATE TRUE\n        WHEN X = 1 DISPLAY \"A\"\n        WHEN X = 2 DISPLAY \"B\"\n        WHEN OTHER DISPLAY \"Z\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["B"]);
}

#[test]
fn evaluate_numeric_when_thru_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 7.",
        "    EVALUATE X\n        WHEN 1 THRU 5 DISPLAY \"L\"\n        WHEN 6 THRU 9 DISPLAY \"H\"\n        WHEN OTHER DISPLAY \"O\"\n    END-EVALUATE.",
    ));
}

#[test]
fn perform_varying_sum_sequence() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 S PIC 9(3) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 4\n        ADD I TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn perform_varying_descending_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 5 BY -1 UNTIL I < 1\n        DISPLAY I\n    END-PERFORM.",
    ));
}

#[test]
fn perform_inline_until_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I > 2\n        ADD 1 TO I\n    END-PERFORM.",
    ));
}

#[test]
fn perform_times_with_nested_if_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM 3 TIMES\n        ADD 1 TO I\n        IF I = 2 DISPLAY \"M\" END-IF\n    END-PERFORM.",
    ));
}

#[test]
fn string_delimited_by_size_and_space() {
    let out = run_prints(&p(
        "01 A PIC X(4) VALUE \"ONE\".\n01 B PIC X(4) VALUE \"TWO\".\n01 R PIC X(20) VALUE SPACES.",
        "    STRING A DELIMITED BY SPACE\n           B DELIMITED BY SPACE\n           INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["ONETWO"]);
}

#[test]
fn unstring_delimited_by_or_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"A,B;C\".\n01 F1 PIC X(3).\n01 F2 PIC X(3).\n01 F3 PIC X(3).",
        "    UNSTRING SRC DELIMITED BY \",\" OR \";\" INTO F1 F2 F3.",
    ));
}

#[test]
fn unstring_delimiter_in_and_count_in_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"AA,BBB\".\n01 F1 PIC X(5).\n01 D1 PIC X.\n01 C1 PIC 9(2).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 DELIMITER IN D1 COUNT IN C1.",
    ));
}

#[test]
fn inspect_tallying_leading_zeroes() {
    let out = run_prints(&p(
        "01 TXT PIC X(8) VALUE \"0001234\".\n01 CNT PIC 9(2) VALUE 0.",
        "    INSPECT TXT TALLYING CNT FOR LEADING \"0\".\n    DISPLAY CNT.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn inspect_replacing_all_letters_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(6) VALUE \"ABCABC\".",
        "    INSPECT TXT REPLACING ALL \"A\" BY \"Z\".",
    ));
}

#[test]
fn search_all_with_at_end_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 TAB.\n   05 E OCCURS 4 TIMES ASCENDING KEY IS K INDEXED BY I.\n      10 K PIC 9(2).\n01 F PIC X VALUE \"N\".\nPROCEDURE DIVISION.\n    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    MOVE 4 TO K(4).\n    SEARCH ALL E\n        AT END MOVE \"N\" TO F\n        WHEN K(I) = 3 MOVE \"Y\" TO F\n    END-SEARCH.\n    DISPLAY F.\n    STOP RUN.",
    );
}

#[test]
fn call_using_by_reference_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUBR\" USING BY REFERENCE X.\n    STOP RUN.",
    );
}

#[test]
fn call_using_by_content_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUBC\" USING BY CONTENT X.\n    STOP RUN.",
    );
}

#[test]
fn call_using_by_value_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUBV\" USING BY VALUE X.\n    STOP RUN.",
    );
}

#[test]
fn stop_run_statement_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn go_to_label_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GO TO A.\nA. STOP RUN.",
    );
}

#[test]
fn alter_statement_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    ALTER A TO PROCEED TO B.\nA. DISPLAY \"A\".\nB. STOP RUN.",
    );
}

#[test]
fn add_corresponding_groups_compiles() {
    compile_ok(&p(
        "01 G1.\n   05 A PIC 9(2) VALUE 11.\n   05 B PIC 9(2) VALUE 22.\n01 G2.\n   05 A PIC 9(2) VALUE 1.\n   05 B PIC 9(2) VALUE 2.",
        "    ADD CORRESPONDING G1 TO G2.",
    ));
}

#[test]
fn add_rounded_compiles() {
    compile_ok(&p(
        "01 A PIC 9V9 VALUE 1.5.\n01 B PIC 9V9 VALUE 2.4.\n01 R PIC 9 VALUE 0.",
        "    ADD A B GIVING R ROUNDED.",
    ));
}

#[test]
fn subtract_corresponding_groups_compiles() {
    compile_ok(&p(
        "01 G1.\n   05 A PIC 9(2) VALUE 9.\n   05 B PIC 9(2) VALUE 8.\n01 G2.\n   05 A PIC 9(2) VALUE 4.\n   05 B PIC 9(2) VALUE 3.",
        "    SUBTRACT CORRESPONDING G2 FROM G1.",
    ));
}

#[test]
fn multiply_rounded_compiles() {
    compile_ok(&p(
        "01 A PIC 9V9 VALUE 1.5.\n01 B PIC 9V9 VALUE 2.5.\n01 R PIC 9 VALUE 0.",
        "    MULTIPLY A BY B GIVING R ROUNDED.",
    ));
}

#[test]
fn divide_remainder_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 17.\n01 B PIC 9(3) VALUE 5.\n01 Q PIC 9(3).\n01 M PIC 9(3).",
        "    DIVIDE B INTO A GIVING Q REMAINDER M.",
    ));
}

#[test]
fn compute_unary_minus_compiles() {
    compile_ok(&p("01 R PIC S9(3) VALUE 0.", "    COMPUTE R = -5 + 2."));
}

#[test]
fn compute_nested_expression_compiles() {
    compile_ok(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = (3 + 5) * (2 + 1).",
    ));
}

#[test]
fn if_next_sentence_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.",
        "    IF A = 1\n        CONTINUE\n    ELSE\n        CONTINUE\n    END-IF.",
    ));
}

#[test]
fn if_class_numeric_compiles() {
    compile_ok(&p(
        "01 X PIC X(5) VALUE \"123\".",
        "    IF X IS NUMERIC DISPLAY \"Y\" END-IF.",
    ));
}

#[test]
fn if_sign_positive_compiles() {
    compile_ok(&p(
        "01 X PIC S9(3) VALUE 3.",
        "    IF X IS POSITIVE DISPLAY \"P\" END-IF.",
    ));
}

#[test]
fn evaluate_multiple_subjects_also_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    EVALUATE A ALSO B\n        WHEN 1 ALSO 2 DISPLAY \"M\"\n        WHEN OTHER DISPLAY \"N\"\n    END-EVALUATE.",
    ));
}

#[test]
fn evaluate_when_any_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 1.",
        "    EVALUATE A ALSO B\n        WHEN 5 ALSO ANY DISPLAY \"HIT\"\n        WHEN OTHER DISPLAY \"MISS\"\n    END-EVALUATE.",
    ));
}

#[test]
fn perform_with_test_before_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM WITH TEST BEFORE UNTIL I > 2\n        ADD 1 TO I\n    END-PERFORM.",
    ));
}

#[test]
fn perform_with_test_after_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM WITH TEST AFTER UNTIL I > 2\n        ADD 1 TO I\n    END-PERFORM.",
    ));
}

#[test]
fn perform_thru_paragraphs_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM P1 THRU P2.\n    STOP RUN.\nP1. DISPLAY \"1\".\nP2. DISPLAY \"2\".",
    );
}

#[test]
fn perform_section_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM S1.\n    STOP RUN.\nS1 SECTION.\nP1. DISPLAY \"S\".",
    );
}

#[test]
fn string_with_pointer_compiles() {
    compile_ok(&p(
        "01 A PIC X(2) VALUE \"AB\".\n01 B PIC X(2) VALUE \"CD\".\n01 R PIC X(10).\n01 P PIC 9(2) VALUE 1.",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R WITH POINTER P.",
    ));
}

#[test]
fn string_on_overflow_compiles() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"ABCDE\".\n01 B PIC X(5) VALUE \"FGHIJ\".\n01 R PIC X(3).",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R ON OVERFLOW DISPLAY \"OV\" END-STRING.",
    ));
}

#[test]
fn unstring_with_pointer_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"AA,BBB,CC\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).\n01 P PIC 9(2) VALUE 1.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 WITH POINTER P.",
    ));
}

#[test]
fn unstring_tallying_in_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(12) VALUE \"A,B,C\".\n01 F1 PIC X(2).\n01 F2 PIC X(2).\n01 F3 PIC X(2).\n01 T PIC 9 VALUE 0.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3 TALLYING IN T.",
    ));
}

#[test]
fn inspect_replacing_leading_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(6) VALUE \"000123\".",
        "    INSPECT TXT REPLACING LEADING \"0\" BY \" \".",
    ));
}

#[test]
fn inspect_converting_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(6) VALUE \"abc123\".",
        "    INSPECT TXT CONVERTING \"abc\" TO \"ABC\".",
    ));
}

#[test]
fn search_linear_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 TAB.\n   05 E OCCURS 3 TIMES INDEXED BY I.\n      10 K PIC 9.\nPROCEDURE DIVISION.\n    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    SET I TO 1.\n    SEARCH E\n        WHEN K(I) = 2 DISPLAY \"Y\"\n    END-SEARCH.\n    STOP RUN.",
    );
}

#[test]
fn search_multiple_when_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 TAB.\n   05 E OCCURS 4 TIMES INDEXED BY I.\n      10 K PIC 9.\nPROCEDURE DIVISION.\n    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    MOVE 4 TO K(4).\n    SET I TO 1.\n    SEARCH E\n        WHEN K(I) = 1 DISPLAY \"A\"\n        WHEN K(I) = 4 DISPLAY \"D\"\n    END-SEARCH.\n    STOP RUN.",
    );
}

#[test]
fn call_returning_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 R PIC 9(3).\nPROCEDURE DIVISION.\n    CALL \"SUBRET\" RETURNING R.\n    STOP RUN.",
    );
}

#[test]
fn call_identifier_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PGM PIC X(8) VALUE \"SUBMOD\".\nPROCEDURE DIVISION.\n    CALL PGM.\n    STOP RUN.",
    );
}

#[test]
fn call_nested_in_if_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 F PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    IF F = 1\n        CALL \"OKM\"\n    ELSE\n        CALL \"NOM\"\n    END-IF.\n    STOP RUN.",
    );
}

#[test]
fn goback_in_subprogram_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. SUB1.\nPROCEDURE DIVISION.\n    GOBACK.");
}

#[test]
fn cancel_identifier_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PGM PIC X(8) VALUE \"SUBMOD\".\nPROCEDURE DIVISION.\n    CANCEL PGM.\n    STOP RUN.",
    );
}

#[test]
fn stop_literal_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    STOP \"DONE\".");
}

#[test]
fn condition_name_set_true_compiles() {
    compile_ok(&p(
        "01 ST PIC X VALUE \"N\".\n   88 ACTIVE VALUE \"Y\".",
        "    SET ACTIVE TO TRUE.",
    ));
}

#[test]
fn condition_name_set_false_compiles() {
    compile_ok(&p(
        "01 ST PIC X VALUE \"Y\".\n   88 ACTIVE VALUE \"Y\".\n   88 INACTIVE VALUE \"N\".",
        "    SET ACTIVE TO FALSE.",
    ));
}

#[test]
fn level_66_renames_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 G.\n   05 A PIC X.\n   05 B PIC X.\n66 AB RENAMES A THRU B.\nPROCEDURE DIVISION.\n    DISPLAY AB.\n    STOP RUN.",
    );
}

#[test]
fn redefines_numeric_alpha_overlay_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N PIC 9(4) VALUE 1234.\n01 N-X REDEFINES N PIC X(4).\nPROCEDURE DIVISION.\n    DISPLAY N-X.\n    STOP RUN.",
    );
}

#[test]
fn move_corresponding_groups_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 G1.\n   05 A PIC 9 VALUE 1.\n   05 B PIC X VALUE \"X\".\n01 G2.\n   05 A PIC 9 VALUE 0.\n   05 B PIC X VALUE \" \".\nPROCEDURE DIVISION.\n    MOVE CORRESPONDING G1 TO G2.\n    STOP RUN.",
    );
}

#[test]
fn initialize_replacing_numeric_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 G.\n   05 A PIC 9 VALUE 5.\n   05 B PIC X VALUE \"Z\".\nPROCEDURE DIVISION.\n    INITIALIZE G REPLACING NUMERIC DATA BY 9.\n    STOP RUN.",
    );
}

#[test]
fn special_names_currency_sign_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CURRENCY SIGN IS \"$\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N PIC $9.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_and_object_computer_with_program_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSOURCE-COMPUTER. IBM-Z.\nOBJECT-COMPUTER. IBM-Z.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}
