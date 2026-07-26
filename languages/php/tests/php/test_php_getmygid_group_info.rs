use super::helpers::run_prints;

#[test]
fn test_getmygid_numeric_or_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$gid = getmygid();
echo ($gid === false || (is_int($gid) && $gid >= 0)) ? 'gid_numeric' : 'err', "\n";
"#
        ),
        vec!["gid_numeric"]
    );
}

#[test]
fn test_getmyuid_numeric_or_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$uid = getmyuid();
echo ($uid === false || (is_int($uid) && $uid >= 0)) ? 'uid_numeric' : 'err', "\n";
"#
        ),
        vec!["uid_numeric"]
    );
}
