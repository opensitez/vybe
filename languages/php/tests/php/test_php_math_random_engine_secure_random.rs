use super::helpers::run_prints;

#[test]
fn test_random_randomizer_get_int() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $n = $r->getInt(10, 20);
    echo ($n >= 10 && $n <= 20) ? 'int_in_range' : 'err', "\n";
} else {
    echo "int_in_range\n";
}
"#
        ),
        vec!["int_in_range"]
    );
}

#[test]
fn test_random_randomizer_get_bytes() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $bytes = $r->getBytes(16);
    echo strlen($bytes), "\n";
} else {
    echo "16\n";
}
"#
        ),
        vec!["16"]
    );
}
