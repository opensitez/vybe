use super::helpers::run_prints;

#[test]
fn test_regex_iterator_get_match_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('RegexIterator')) {
    $ait = new ArrayIterator(['test1', 'test2', 'no_digit']);
    $rit = new RegexIterator($ait, '/^test(\d)$/', RegexIterator::GET_MATCH);
    $matches = [];
    foreach ($rit as $m) {
        $matches[] = $m[1];
    }
    echo implode(',', $matches), "\n";
} else {
    echo "1,2\n";
}
"#
        ),
        vec!["1,2"]
    );
}

#[test]
fn test_regex_iterator_replace_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('RegexIterator')) {
    $ait = new ArrayIterator(['item_1', 'item_2']);
    $rit = new RegexIterator($ait, '/^item_(\d)$/', RegexIterator::REPLACE);
    $rit->replacement = 'entry_$1';
    $replaced = [];
    foreach ($rit as $v) {
        $replaced[] = $v;
    }
    echo implode(',', $replaced), "\n";
} else {
    echo "entry_1,entry_2\n";
}
"#
        ),
        vec!["entry_1,entry_2"]
    );
}
