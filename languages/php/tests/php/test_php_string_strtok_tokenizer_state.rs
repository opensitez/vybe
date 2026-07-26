use super::helpers::run_prints;

#[test]
fn test_strtok_basic_sequence() {
    assert_eq!(
        run_prints(
            r#"<?php
$string = "hello,world;php/test";
$tok = strtok($string, ",;/");
$tokens = [];
while ($tok !== false) {
    $tokens[] = $tok;
    $tok = strtok(",;/");
}
echo implode('-', $tokens), "\n";
"#
        ),
        vec!["hello-world-php-test"]
    );
}

#[test]
fn test_strtok_reinitialization() {
    assert_eq!(
        run_prints(
            r#"<?php
strtok("first token", " ");
$first = strtok("second token", " ");
echo $first, "\n";
"#
        ),
        vec!["second"]
    );
}

#[test]
fn test_strtok_reset_by_empty_subject() {
    assert_eq!(
        run_prints(
            r#"<?php
strtok("a,b", ",");
echo strtok(",", "x"), "\n";
echo strtok("", ",") === false ? 'reset' : 'not_reset';
"#
        ),
        vec!["b|reset"]
    );
}
