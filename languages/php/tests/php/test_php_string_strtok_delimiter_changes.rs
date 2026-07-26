use super::helpers::run_prints;

#[test]
fn test_strtok_change_delimiters_mid_loop() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "name=John&age=30";
$key = strtok($str, "=");
$val = strtok("&");
echo $key . ':' . $val, "\n";
"#
        ),
        vec!["name:John"]
    );
}

#[test]
fn test_strtok_false_on_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$tok = strtok("", " ");
echo $tok === false ? 'false_on_empty' : 'err', "\n";
"#
        ),
        vec!["false_on_empty"]
    );
}

#[test]
fn test_strtok_multiple_calls_with_dynamic_delimiters() {
    assert_eq!(
        run_prints(
            r#"<?php
$tok = strtok("a,b;c,d", ",");
echo $tok . "|";
$tok = strtok(" ;");
echo $tok . "|";
$tok = strtok(" ,;");
echo $tok;
"#
        ),
        vec!["a|b|c"]
    );
}

#[test]
fn test_strtok_with_unicode_and_multidelim() {
    assert_eq!(
        run_prints(
            r#"<?php
$tok = strtok("x|y;z", "|;");
echo $tok;
echo "|";
$tok = strtok(" ;");
echo $tok;
echo "|";
$tok = strtok(";");
echo $tok;
"#
        ),
        vec!["x|y|z"]
    );
}
