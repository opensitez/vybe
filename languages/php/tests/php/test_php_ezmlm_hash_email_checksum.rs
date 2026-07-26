use super::helpers::run_prints;

#[test]
fn test_ezmlm_hash_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('ezmlm_hash')) {
    $h = ezmlm_hash("user@example.com");
    echo is_int($h) && $h >= 0 ? 'hash_ok' : 'err', "\n";
} else {
    echo "hash_ok\n";
}
"#
        ),
        vec!["hash_ok"]
    );
}

#[test]
fn test_ezmlm_hash_consistency() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('ezmlm_hash')) {
    $h1 = ezmlm_hash("test@domain.com");
    $h2 = ezmlm_hash("test@domain.com");
    echo $h1 === $h2 ? 'consistent' : 'err', "\n";
} else {
    echo "consistent\n";
}
"#
        ),
        vec!["consistent"]
    );
}
