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

fn int_sequence(len: i64) -> Vec<i64> {
    (1..=len).collect()
}

#[test]
fn php_string_manipulation() {
    for len in 1..=20_i64 {
        let values = int_sequence(len)
            .iter()
            .map(|value| value.to_string())
            .collect::<Vec<_>>()
            .join(",");

        let arr = format!("[{}]", values);
        let joined = (1..=len).map(|value| value.to_string()).collect::<Vec<_>>().join("-");
        let joined_comma = (1..=len).map(|value| value.to_string()).collect::<Vec<_>>().join(",");
        let exploded_count = len;
        let dash_len = joined.len() as i64;

        assert_output(
            &format!("implode(',', {arr})"),
            &joined_comma,
        );
        assert_int(
            &format!("count(explode('-', {}))", quote_php(&joined)),
            exploded_count,
        );
        assert_int(&format!("strlen(implode('-', {arr}))"), dash_len);
        assert_output(
            &format!("str_replace('-', ':', {})", quote_php(&joined)),
            &joined_comma,
        );
        assert_int(
            &format!("strlen(str_replace(' ', '-', trim({})))", quote_php(&format!("  {joined}  "))),
            dash_len,
        );
        assert_output(
            &format!("substr(implode('-', {arr}), 0, 3)"),
            &joined.chars().take(3).collect::<String>(),
        );
    }
}

#[test]
fn php_string_split_and_join_with_empty_parts() {
    let out = run_prints(
        "<?php\n$source = ',a,,b,';\necho count(explode(',', $source)) . '|';\necho implode('|', explode(',', $source));\n",
    );
    assert_eq!(out, vec!["5| |a||b|"]);
}

#[test]
fn php_string_explode_limit_keeps_remainder() {
    let out = run_prints(
        "<?php\n$parts = explode('-', 'a-b-c-d', 3);\nforeach ($parts as $part) { echo $part; }\n",
    );
    assert_eq!(out, vec!["abc-d"]);
}

#[test]
fn php_string_implode_without_glue() {
    let out = run_prints(
        "<?php\n$a = ['php', '7', 'test'];\necho implode($a);\n",
    );
    assert_eq!(out, vec!["php7test"]);
}

#[test]
fn php_string_implode_handles_nested_arrays() {
    let out = run_prints(
        "<?php\n$a = ['a', ['x']];\ntry {\n    echo implode('-', $a);\n} catch (TypeError $e) {\n    echo 'type_error';\n}\n",
    );
    assert_eq!(out, vec!["type_error"]);
}

#[test]
fn php_string_list_and_join_chain() {
    let out = run_prints(
        "<?php\n$parts = ['  a ', ' b ', 'c '];\n$normalized = array_map('trim', $parts);\n$joined = implode(':', $normalized);\n$tokens = explode(':', $joined);\necho $tokens[0] . '|' . $tokens[2];\n",
    );
    assert_eq!(out, vec!["a|c"]);
}

#[test]
fn php_string_parse_words_with_strtok() {
    let out = run_prints(
        "<?php\n$str = 'red,green,blue';\n$first = strtok($str, ',');\n$second = strtok(',');\n$third = strtok(',');\necho $first . '|' . $second . '|' . $third;\n",
    );
    assert_eq!(out, vec!["red|green|blue"]);
}
