use super::helpers::run_prints;

#[test]
fn test_session_cache_limiter_getter_setter() {
    assert_eq!(
        run_prints(
            r#"<?php
$old = session_cache_limiter('private');
echo session_cache_limiter() . '|' . (is_string($old) ? 'old_limiter_ok' : 'err'), "\n";
"#
        ),
        vec!["private|old_limiter_ok"]
    );
}

#[test]
fn test_session_cache_expire_getter_setter() {
    assert_eq!(
        run_prints(
            r#"<?php
$old = session_cache_expire(60);
echo session_cache_expire() . '|' . (is_int($old) ? 'old_expire_ok' : 'err'), "\n";
"#
        ),
        vec!["60|old_expire_ok"]
    );
}
