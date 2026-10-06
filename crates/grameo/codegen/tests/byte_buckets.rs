use grameo_codegen::generate;

#[test]
fn mixed_ascii_case_and_unicode_first_bytes_keep_priority_and_grouping() {
    let output = generate(r#"root = { (^"a" | ^"A" | ^"é" | ^"ß" | ^"1" | "é" | "a" | "!" | ^"b" | "Z" | ^"Z" | "abc") ~ EOI }"#).unwrap();
    for arm in [
        "Some(65) => &[0, 1]",
        "Some(97) => &[0, 1, 6, 11]",
        "Some(195) => &[2, 3, 5]",
        "Some(90) => &[9, 10]",
        "Some(122) => &[10]",
    ] {
        assert!(output.contains(arm), "missing ordered candidate arm {arm}");
    }
}

#[test]
fn printable_ascii_and_case_sensitive_duplicates_keep_original_priority() {
    let choices = (32u8..=126)
        .flat_map(|byte| {
            let text = char::from(byte).to_string();
            [format!("{text:?}"), format!("^{text:?}")]
        })
        .collect::<Vec<_>>()
        .join(" | ");
    let output = generate(&format!("root = {{ ({choices}) ~ EOI }}")).unwrap();
    for arm in [
        "Some(34) => &[4, 5]",
        "Some(65) => &[66, 67, 131]",
        "Some(97) => &[67, 130, 131]",
        "Some(92) => &[120, 121]",
    ] {
        assert!(output.contains(arm), "missing ordered candidate arm {arm}");
    }
}

#[test]
fn empty_alternative_retains_candidate_dispatch_fallback() {
    let output = generate(r#"root = { ("" | "a" | ^"a" | "b" | "c" | "d" | "e" | "f" | "g" | "h") ~ EOI }"#).unwrap();
    let candidates = output.split("fn terminal_candidates").nth(1).unwrap()
        .split("fn failure_terminals").next().unwrap();
    assert!(!candidates.contains("Some(match _byte"));
}
