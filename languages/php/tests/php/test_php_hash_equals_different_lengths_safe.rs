use super::helpers::run_prints;

#[test]
fn test_hash_equals_unequal_length_false() {
    assert_eq!(
        run_prints(
            r#"<?php
echo hash_equals('short', 'much_longer_string') ? 'equal' : 'not_equal', "\n";
"#
        ),
        vec!["not_equal"]
    );
}

#[test]
fn test_hash_equals_empty_strings() {
    assert_eq!(
        run_prints(
            r#"<?php
echo hash_equals('', '') ? 'equal_empty' : 'not_equal', "\n";
"#
        ),
        vec!["equal_empty"]
    );
}
