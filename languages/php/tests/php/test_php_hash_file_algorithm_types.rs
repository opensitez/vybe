use super::helpers::run_prints;

#[test]
fn test_hash_file_sha256_digest() {
    assert_eq!(
        run_prints(
            r#"<?php
$tmp = sys_get_temp_dir() . '/test_hash_file.txt';
file_put_contents($tmp, 'hello world');
$digest = hash_file('sha256', $tmp);
unlink($tmp);
echo strlen($digest), "\n";
"#
        ),
        vec!["64"]
    );
}

#[test]
fn test_hash_file_raw_output() {
    assert_eq!(
        run_prints(
            r#"<?php
$tmp = sys_get_temp_dir() . '/test_hash_file_raw.txt';
file_put_contents($tmp, 'data');
$raw = hash_file('md5', $tmp, true);
unlink($tmp);
echo strlen($raw), "\n";
"#
        ),
        vec!["16"]
    );
}
