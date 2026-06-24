use crate::helpers::run_main;

#[test]
fn string_joiner_add_two_elements_with_delimiter() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(", "); sj.add("a"); sj.add("b"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["a, b"]);
}

#[test]
fn string_joiner_add_single_element_has_no_delimiter() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner("-"); sj.add("solo"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn string_joiner_empty_after_construction_uses_empty_value() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner("|", "[", "]"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn string_joiner_set_empty_value_customizes_empty_representation() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(",", "(", ")"); sj.setEmptyValue("none"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn string_joiner_prefix_and_suffix_wrap_joined_elements() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(",", "<", ">"); sj.add("x"); sj.add("y"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["<x,y>"]);
}

#[test]
fn string_joiner_add_three_elements_preserves_order() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(":"); sj.add("one"); sj.add("two"); sj.add("three"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["one:two:three"]);
}

#[test]
fn string_joiner_merge_combines_non_empty_joiner() {
    let out = run_main(
        r#"java.util.StringJoiner left = new java.util.StringJoiner(", "); left.add("a"); java.util.StringJoiner right = new java.util.StringJoiner(", "); right.add("b"); left.merge(right); System.out.println(left.toString());"#,
    );
    assert_eq!(out, vec!["a, b"]);
}

#[test]
fn string_joiner_merge_empty_other_joiner_is_noop() {
    let out = run_main(
        r#"java.util.StringJoiner left = new java.util.StringJoiner("|"); left.add("keep"); java.util.StringJoiner right = new java.util.StringJoiner("|"); left.merge(right); System.out.println(left.toString());"#,
    );
    assert_eq!(out, vec!["keep"]);
}

#[test]
fn string_joiner_merge_into_empty_adopts_other_content() {
    let out = run_main(
        r#"java.util.StringJoiner left = new java.util.StringJoiner("-"); java.util.StringJoiner right = new java.util.StringJoiner("-"); right.add("z"); left.merge(right); System.out.println(left.toString());"#,
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn string_joiner_length_counts_characters_in_current_value() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(", "); sj.add("ab"); sj.add("c"); System.out.println(sj.length());"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn string_joiner_length_of_empty_joiner_counts_prefix_and_suffix() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(",", "(", ")"); System.out.println(sj.length());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn string_join_varargs_joins_three_strings_with_hyphen() {
    let out = run_main(r#"System.out.println(String.join("-", "a", "b", "c"));"#);
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn string_join_varargs_single_element_has_no_delimiter() {
    let out = run_main(r#"System.out.println(String.join("-", "solo"));"#);
    assert_eq!(out, vec!["solo"]);
}

#[test]
fn string_join_empty_string_array_yields_empty_string() {
    let out = run_main(r#"System.out.println(String.join(",", new String[] {}));"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn string_join_char_sequence_delimiter_with_two_elements() {
    let out = run_main(r#"System.out.println(String.join("::", "first", "second"));"#);
    assert_eq!(out, vec!["first::second"]);
}

#[test]
fn string_join_from_array_literal() {
    let out = run_main(
        r#"String[] parts = {"red", "green", "blue"}; System.out.println(String.join("/", parts));"#,
    );
    assert_eq!(out, vec!["red/green/blue"]);
}

#[test]
fn string_join_iterable_from_arraylist() {
    let out = run_main(
        "java.util.ArrayList<String> items = new java.util.ArrayList<String>(); items.add(\"x\"); items.add(\"y\"); System.out.println(String.join(\"+\", items));",
    );
    assert_eq!(out, vec!["x+y"]);
}

#[test]
fn string_joiner_add_returns_same_instance_for_chaining() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(""); sj.add("a").add("b").add("c"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn string_joiner_with_pipe_delimiter_and_brackets() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner("|", "{", "}"); sj.add("1"); sj.add("2"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["{1|2}"]);
}

#[test]
fn string_joiner_merge_preserves_other_delimiter_formatting() {
    let out = run_main(
        r#"java.util.StringJoiner a = new java.util.StringJoiner(" + "); a.add("1"); java.util.StringJoiner b = new java.util.StringJoiner(" + "); b.add("2"); b.add("3"); a.merge(b); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["1 + 2 + 3"]);
}

#[test]
fn string_join_two_empty_strings_with_comma() {
    let out = run_main(r#"System.out.println(String.join(",", "", ""));"#);
    assert_eq!(out, vec![","]);
}

#[test]
fn string_join_numeric_strings_as_text() {
    let out = run_main(r#"System.out.println(String.join("", "1", "2", "3"));"#);
    assert_eq!(out, vec!["123"]);
}

#[test]
fn string_joiner_add_after_set_empty_value_shows_content() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(",", "(", ")"); sj.setEmptyValue("empty"); sj.add("hi"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["(hi)"]);
}

#[test]
fn string_joiner_merge_both_with_prefix_suffix() {
    let out = run_main(
        r#"java.util.StringJoiner a = new java.util.StringJoiner(",", "[", "]"); a.add("a"); java.util.StringJoiner b = new java.util.StringJoiner(",", "(", ")"); b.add("b"); a.merge(b); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["[a,(b)]"]);
}

#[test]
fn string_join_whitespace_delimiter_between_words() {
    let out = run_main(r#"System.out.println(String.join(" ", "hello", "world"));"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn string_joiner_four_elements_with_slash() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner("/"); sj.add("2024"); sj.add("06"); sj.add("24"); sj.add("vybe"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["2024/06/24/vybe"]);
}

#[test]
fn string_join_array_with_one_element() {
    let out = run_main(
        r#"String[] only = {"one"}; System.out.println(String.join("-", only));"#,
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn string_joiner_empty_with_custom_empty_value_and_wrappers() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(";", "<", ">"); sj.setEmptyValue("<empty>"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["<empty>"]);
}

#[test]
fn string_join_repeated_delimiter_characters() {
    let out = run_main(r#"System.out.println(String.join("...", "wait", "for", "it"));"#);
    assert_eq!(out, vec!["wait...for...it"]);
}

#[test]
fn string_joiner_add_numeric_strings() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(""); sj.add("10"); sj.add("20"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["1020"]);
}

#[test]
fn string_join_char_sequence_from_string_builder_delimiter() {
    let out = run_main(
        r#"StringBuilder delim = new StringBuilder("->"); System.out.println(String.join(delim.toString(), "go", "stop"));"#,
    );
    assert_eq!(out, vec!["go->stop"]);
}

#[test]
fn string_joiner_length_after_merge_includes_delimiter() {
    let out = run_main(
        r#"java.util.StringJoiner a = new java.util.StringJoiner(","); a.add("aa"); java.util.StringJoiner b = new java.util.StringJoiner(","); b.add("bbb"); a.merge(b); System.out.println(a.length());"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn string_join_five_varargs_elements() {
    let out = run_main(r#"System.out.println(String.join("", "a", "b", "c", "d", "e"));"#);
    assert_eq!(out, vec!["abcde"]);
}

#[test]
fn string_joiner_merge_chain_three_joiners() {
    let out = run_main(
        r#"java.util.StringJoiner a = new java.util.StringJoiner(","); a.add("1"); java.util.StringJoiner b = new java.util.StringJoiner(","); b.add("2"); java.util.StringJoiner c = new java.util.StringJoiner(","); c.add("3"); a.merge(b); a.merge(c); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn string_join_empty_delimiter_concatenates_directly() {
    let out = run_main(r#"System.out.println(String.join("", "foo", "bar"));"#);
    assert_eq!(out, vec!["foobar"]);
}

#[test]
fn string_joiner_to_string_matches_manual_join_for_two_items() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner(" | "); sj.add("alpha"); sj.add("beta"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["alpha | beta"]);
}

#[test]
fn string_join_arraylist_three_items_with_colon() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"r\"); list.add(\"g\"); list.add(\"b\"); System.out.println(String.join(\":\", list));",
    );
    assert_eq!(out, vec!["r:g:b"]);
}

#[test]
fn string_joiner_set_empty_value_then_merge_still_joins() {
    let out = run_main(
        r#"java.util.StringJoiner a = new java.util.StringJoiner(",", "(", ")"); a.setEmptyValue("void"); java.util.StringJoiner b = new java.util.StringJoiner(","); b.add("z"); a.merge(b); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["(z)"]);
}

#[test]
fn string_join_two_element_array_with_star_delimiter() {
    let out = run_main(
        r#"String[] pair = {"left", "right"}; System.out.println(String.join("*", pair));"#,
    );
    assert_eq!(out, vec!["left*right"]);
}

#[test]
fn string_joiner_add_empty_string_element() {
    let out = run_main(
        r#"java.util.StringJoiner sj = new java.util.StringJoiner("-"); sj.add(""); sj.add("end"); System.out.println(sj.toString());"#,
    );
    assert_eq!(out, vec!["-end"]);
}
