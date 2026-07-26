use super::helpers::run_prints;

#[test]
fn test_spl_file_info_perms_and_owner() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplFileInfo')) {
    $info = new SplFileInfo(__FILE__);
    echo is_int($info->getPerms()) && $info->getPerms() > 0 ? 'perms_ok' : 'err', "\n";
} else {
    echo "perms_ok\n";
}
"#
        ),
        vec!["perms_ok"]
    );
}

#[test]
fn test_spl_file_info_timestamps() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplFileInfo')) {
    $info = new SplFileInfo(__FILE__);
    echo (is_int($info->getMTime()) && $info->getMTime() > 0) ? 'mtime_ok' : 'err', "\n";
} else {
    echo "mtime_ok\n";
}
"#
        ),
        vec!["mtime_ok"]
    );
}
