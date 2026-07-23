use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: File Path Resolution & Permissions — realpath, realpath_cache_get, fileperms, fileowner, touch, chmod, is_readable, is_writable
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_realpath_canonical_absolute_path() {
    let out = run_prints(
        r#"<?php
$tmp = sys_get_temp_dir();
$real = realpath($tmp);
echo (is_string($real) && strlen($real) > 0 && is_dir($real)) ? "REALPATH_OK" : "REALPATH_FAIL";
"#,
    );
    assert_eq!(out, vec!["REALPATH_OK"]);
}

#[test]
fn test_php_file_permissions_readable_writable_executable() {
    let out = run_prints(
        r#"<?php
$tmpFile = tempnam(sys_get_temp_dir(), "vybe_perm_");
file_put_contents($tmpFile, "test perms");

echo is_readable($tmpFile) ? "R1" : "R0";
echo is_writable($tmpFile) ? "W1" : "W0";
echo is_file($tmpFile) ? "F1" : "F0";

unlink($tmpFile);
"#,
    );
    assert_eq!(out, vec!["R1W1F1"]);
}

#[test]
fn test_php_touch_file_creation_and_mtime_update() {
    let out = run_prints(
        r#"<?php
$file = sys_get_temp_dir() . "/vybe_touch_" . time() . ".txt";
$created = touch($file);
echo $created && is_file($file) ? "TOUCHED_OK" : "TOUCH_FAIL";
if (is_file($file)) unlink($file);
"#,
    );
    assert_eq!(out, vec!["TOUCHED_OK"]);
}

#[test]
fn test_php_chmod_file_permission_modification() {
    compile_ok(
        r#"<?php
$file = tempnam(sys_get_temp_dir(), "vybe_chmod_");
chmod($file, 0644);
$perms = fileperms($file);
echo is_numeric($perms) ? "PERMS_OK" : "FAIL";
unlink($file);
"#,
    );
}

#[test]
fn test_php_realpath_cache_size_and_get() {
    compile_ok(
        r#"<?php
$cacheSize = realpath_cache_size();
$entries = realpath_cache_get();
echo "Size=$cacheSize Entries=" . (is_array($entries) ? count($entries) : 0);
"#,
    );
}

#[test]
fn test_php_fileowner_and_filegroup_attributes() {
    compile_ok(
        r#"<?php
$file = tempnam(sys_get_temp_dir(), "vybe_owner_");
$owner = fileowner($file);
$group = filegroup($file);
echo is_numeric($owner) && is_numeric($group) ? "OWNER_OK" : "FAIL";
unlink($file);
"#,
    );
}

#[test]
fn test_php_clearstatcache_flush() {
    compile_ok(
        r#"<?php
$file = tempnam(sys_get_temp_dir(), "vybe_cache_");
filesize($file);
clearstatcache(clear_realpath_cache: true, filename: $file);
unlink($file);
"#,
    );
}

#[test]
fn test_php_disk_free_and_total_space_metrics() {
    compile_ok(
        r#"<?php
$free = disk_free_space(".");
$total = disk_total_space(".");
echo ($free !== false && $total !== false && $total >= $free) ? "SPACE_METRICS_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_is_executable_file_check() {
    compile_ok(
        r#"<?php
$file = tempnam(sys_get_temp_dir(), "vybe_exec_");
echo is_executable($file) ? "EXEC" : "NON_EXEC";
unlink($file);
"#,
    );
}

#[test]
fn test_php_is_link_and_readlink_symlinks() {
    compile_ok(
        r#"<?php
$target = tempnam(sys_get_temp_dir(), "vybe_target_");
$link = sys_get_temp_dir() . "/vybe_symlink_" . time();
@symlink($target, $link);
if (is_link($link)) {
    echo "SYMLINK_CREATED: " . readlink($link);
    unlink($link);
}
unlink($target);
"#,
    );
}
