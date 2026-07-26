use super::helpers::run_prints;

#[test]
fn test_getlastmod_returns_timestamp_or_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$mod = getlastmod();
echo ($mod === false || (is_int($mod) && $mod > 0)) ? 'lastmod_ok' : 'err', "\n";
"#
        ),
        vec!["lastmod_ok"]
    );
}

#[test]
fn test_getmygid_returns_int_or_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$gid = getmygid();
echo ($gid === false || (is_int($gid) && $gid >= 0)) ? 'gid_ok' : 'err', "\n";
"#
        ),
        vec!["gid_ok"]
    );
}

#[test]
fn test_getmyinode_returns_int_or_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$inode = getmyinode();
echo ($inode === false || (is_int($inode) && $inode >= 0)) ? 'inode_ok' : 'err', "\n";
"#
        ),
        vec!["inode_ok"]
    );
}
