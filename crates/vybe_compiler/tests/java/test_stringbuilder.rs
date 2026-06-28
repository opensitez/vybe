use crate::helpers::run_main;

#[test]
fn stringbuilder_empty_tostring_is_empty_string() {
    let out =
        run_main(r#"StringBuilder sb = new StringBuilder(); System.out.println(sb.toString());"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn stringbuilder_append_single_character() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("x"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn stringbuilder_append_chain_builds_word() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("j").append("a").append("v").append("a"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn stringbuilder_append_to_existing_seed_string() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("pre"); sb.append("-"); sb.append("fix"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["pre-fix"]);
}

#[test]
fn stringbuilder_append_integer_coerces_to_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append(42); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn stringbuilder_append_boolean_coerces_to_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append(true); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stringbuilder_insert_puts_text_at_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ace"); sb.insert(1, "b"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abce"]);
}

#[test]
fn stringbuilder_insert_at_start_prefixes_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("end"); sb.insert(0, "start-"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["start-end"]);
}

#[test]
fn stringbuilder_insert_at_end_appends_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ab"); sb.insert(2, "cd"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abcd"]);
}

#[test]
fn stringbuilder_delete_removes_subrange() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.delete(1, 4); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["adf"]);
}

#[test]
fn stringbuilder_delete_entire_buffer() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abc"); sb.delete(0, 3); System.out.println(sb.toString().length());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_delete_char_at_removes_single_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcde"); sb.deleteCharAt(2); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abde"]);
}

#[test]
fn stringbuilder_reverse_flips_characters() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("stressed"); sb.reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["desserts"]);
}

#[test]
fn stringbuilder_reverse_twice_restores_original() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("pal"); sb.reverse(); sb.reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["pal"]);
}

#[test]
fn stringbuilder_tostring_after_mutations_returns_current_buffer() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("a"); sb.append("b"); sb.append("c"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn stringbuilder_length_reports_character_count() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("hello"); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stringbuilder_length_zero_on_empty_builder() {
    let out =
        run_main(r#"StringBuilder sb = new StringBuilder(); System.out.println(sb.length());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_capacity_constructor_accepts_initial_size() {
    let out =
        run_main(r#"StringBuilder sb = new StringBuilder(64); System.out.println(sb.length());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_capacity_constructor_then_append_preserves_content() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(32); sb.append("seed"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["seed"]);
}

#[test]
fn stringbuilder_chained_append_insert_reverse() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ab"); sb.append("cd").insert(2, "-").reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["dc-ba"]);
}

#[test]
fn stringbuilder_append_returns_same_instance_for_chaining() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("a").append("b"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn stringbuilder_insert_then_append_extends_buffer() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("x"); sb.insert(1, "y"); sb.append("z"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["xyz"]);
}

#[test]
fn stringbuilder_delete_then_append_replaces_removed_span() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.delete(2, 5); sb.append("Z"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abZ"]);
}

#[test]
fn stringbuilder_append_null_literal_prints_null_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append((String) null); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn stringbuilder_append_multiple_lines_with_newline_char() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("line1\n"); sb.append("line2"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["line1\nline2"]);
}

#[test]
fn stringbuilder_reverse_single_character_is_unchanged() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("z"); sb.reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn stringbuilder_insert_empty_string_is_noop_on_content() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ok"); sb.insert(1, ""); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn stringbuilder_delete_char_at_first_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("zap"); sb.deleteCharAt(0); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["ap"]);
}

#[test]
fn stringbuilder_delete_char_at_last_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("zap"); sb.deleteCharAt(2); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["za"]);
}

#[test]
fn stringbuilder_append_doubles_via_chained_calls() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("("); sb.append("inner").append(")"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["(inner)"]);
}

#[test]
fn stringbuilder_length_grows_after_append() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("abc"); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_length_shrinks_after_delete() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.delete(1, 4); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_tostring_matches_length_for_known_content() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("vybe"); System.out.println(sb.toString().length()); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn stringbuilder_mixed_insert_delete_reverse_pipeline() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ab"); sb.insert(2, "c"); sb.delete(0, 1); sb.reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["cba"]);
}

#[test]
fn stringbuilder_seed_constructor_plus_append_builds_sentence() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("Hello"); sb.append(", "); sb.append("world"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["Hello, world"]);
}
