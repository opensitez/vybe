use super::helpers::run_prints;

fn assert_output(expr: &str, expected: &str) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

fn assert_int(expr: &str, expected: i64) {
    assert_output(expr, &expected.to_string());
}

fn quote_php(value: &str) -> String {
    format!("'{}'", value.replace('\\', "\\\\").replace('\'', "\\'"))
}

#[test]
fn php_mbstring_ascii_framework_strings() {
    let words: [&str; 13] = [
        "Framework",
        "Symfony",
        "Laravel",
        "WordPress",
        "Middleware",
        "Repository",
        "Controller",
        "Collection",
        "Validation",
        "Session",
        "Exception",
        "Routing",
        "Provider",
    ];

    for text in words {
        let lower = text.to_lowercase();
        let upper = text.to_uppercase();
        let len = text.chars().count();
        let chunks = [1_i64, 2_i64, 3_i64, 4_i64];

        assert_output(&format!("mb_strtolower({}, 'UTF-8')", quote_php(text)), &lower);
        assert_output(&format!("mb_strtoupper({}, 'UTF-8')", quote_php(text)), &upper);
        assert_int(&format!("mb_strlen({}, 'UTF-8')", quote_php(text)), len as i64);

        // Keep token checks stable on ASCII-only inputs.
        if text.contains('a') {
            let first_a = text.find('a').expect("token exists") as i64;
            let last_a = text.rfind('a').expect("token exists") as i64;
            assert_int(
                &format!("mb_strpos({}, 'a', 0, 'UTF-8')", quote_php(text)),
                first_a,
            );
            assert_int(
                &format!("mb_strrpos({}, 'a', 0, 'UTF-8')", quote_php(text)),
                last_a,
            );
            let count_a = text.matches('a').count() as i64;
            assert_int(
                &format!("mb_substr_count({}, 'a', 'UTF-8')", quote_php(text)),
                count_a,
            );
        }

        for chunk in chunks {
            let width = text.chars().count() as f64 / chunk as f64;
            let split_count = width.ceil() as i64;
            assert_int(
                &format!("count(mb_str_split({}, {}, 'UTF-8'))", quote_php(text), chunk),
                split_count,
            );
        }
    }
}

#[test]
fn php_mbstring_unicode() {
    let unicode_words: [&str; 8] = [
        "世界",
        "ありがとう",
        "こんにちは",
        "über",
        "naïve",
        "café",
        "Français",
        "ユーザー",
    ];

    for text in unicode_words.iter() {
        let len = text.chars().count() as i64;
        assert_int(&format!("mb_strlen({}, 'UTF-8')", quote_php(text)), len);
        assert_int(
            &format!("mb_strlen(mb_substr({}, 0, 1, 'UTF-8'), 'UTF-8')", quote_php(text)),
            1,
        );

        // Ensure multibyte split counts stay consistent with UTF-8 chunking expectations.
        for chunk in [1_i64, 2_i64] {
            let split_count = ((len + chunk - 1) / chunk) as i64;
            assert_int(
                &format!("count(mb_str_split({}, {}, 'UTF-8'))", quote_php(text), chunk),
                split_count,
            );
        }

        // Roundtrip convert across UTF-8 should preserve byte length for text literals.
        assert_int(
            &format!(
                "strlen(mb_convert_encoding(mb_convert_encoding({}, 'UTF-8', 'UTF-8'), 'UTF-8', 'UTF-8'))",
                quote_php(text),
            ),
            text.len() as i64,
        );

        // Indexing stable across common offsets for short UTF-8 strings.
        if len >= 2 {
            let needle = text.chars().next().expect("non-empty")
                .to_string();
            let next = text.chars().nth(1).expect("non-empty");
            assert_int(
                &format!("mb_strpos({}, '{}', 0, 'UTF-8')", quote_php(text), needle),
                0,
            );
            assert_int(
                &format!(
                    "mb_strpos({}, '{}', 0, 'UTF-8')",
                    quote_php(text),
                    next
                ),
                1,
            );
        }

        // Framework-style case folding checks on multibyte content where mapping is stable.
        assert_int(
            &format!("strlen(mb_convert_case({}, MB_CASE_UPPER, 'UTF-8'))", quote_php(text)),
            len,
        );
    }
}

#[test]
fn php_mbstring_unicode_edge_cases_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_check_encoding("Hello", ["ASCII", "UTF-8"]) !== false ? "ok" : "bad";
echo "|";
echo mb_stripos("Café", "CAFÉ", 0, "UTF-8");
echo "|";
echo mb_str_split("😀", 3, "UTF-8")[0] === "😀" ? "one" : "no";
"#,
    );
    assert_eq!(out, vec!["ok|0|one"]);
}

#[test]
fn php_mbstring_unicode_length_consistency_runtime() {
    let out = run_prints(
        r#"<?php
$s = "naïve";
echo mb_strlen($s, "UTF-8");
echo "|";
echo strlen(mb_convert_encoding($s, "UTF-8", "UTF-8"));
"#,
    );
    assert_eq!(out, vec!["5|6"]);
}
