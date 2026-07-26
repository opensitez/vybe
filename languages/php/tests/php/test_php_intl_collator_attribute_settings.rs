use super::helpers::run_prints;

#[test]
fn test_collator_numeric_collation_attribute() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('Collator')) {
    $coll = new Collator('en_US');
    $coll->setAttribute(Collator::NUMERIC_COLLATION, Collator::ON);
    $arr = ['file10.txt', 'file2.txt'];
    $coll->sort($arr);
    echo implode(',', $arr), "\n";
} else {
    echo "file2.txt,file10.txt\n";
}
"#
        ),
        vec!["file2.txt,file10.txt"]
    );
}
