use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Streams & Filesystem Wrapper Operations — file_get_contents, file_put_contents, fopen, fread, fwrite, pathinfo, basename, dirname, realpath
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_pathinfo_filename_extension_directory() {
    let out = run_prints(
        r#"<?php
$path = "/var/www/html/index.controller.php";
$info = pathinfo($path);
echo $info["dirname"] . " | " . $info["basename"] . " | " . $info["extension"] . " | " . $info["filename"];
"#,
    );
    assert_eq!(
        out,
        vec!["/var/www/html | index.controller.php | php | index.controller"]
    );
}

#[test]
fn test_php_basename_dirname_suffix_strip() {
    let out = run_prints(
        r#"<?php
$path = "/home/user/documents/report.pdf";
echo basename($path, ".pdf") . " in " . dirname($path);
"#,
    );
    assert_eq!(out, vec!["report in /home/user/documents"]);
}

#[test]
fn test_php_temp_file_creation_write_read() {
    let out = run_prints(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "vybe_test_");
file_put_contents($tmp, "Hello Filesystem");
echo file_get_contents($tmp);
unlink($tmp);
"#,
    );
    assert_eq!(out, vec!["Hello Filesystem"]);
}

#[test]
fn test_php_file_append_and_lock_flags() {
    let out = run_prints(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "vybe_append_");
file_put_contents($tmp, "Line 1\n");
file_put_contents($tmp, "Line 2\n", FILE_APPEND | LOCK_EX);
echo implode("-", file($tmp, FILE_IGNORE_NEW_LINES));
unlink($tmp);
"#,
    );
    assert_eq!(out, vec!["Line 1-Line 2"]);
}

#[test]
fn test_php_fopen_fread_fwrite_fclose_handle() {
    let out = run_prints(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "vybe_handle_");
$h = fopen($tmp, "w+");
fwrite($h, "Stream Data");
rewind($h);
echo fread($h, 1024);
fclose($h);
unlink($tmp);
"#,
    );
    assert_eq!(out, vec!["Stream Data"]);
}

#[test]
fn test_php_mkdir_rmdir_recursive_directory() {
    compile_ok(
        r#"<?php
$dir = sys_get_temp_dir() . "/nested/dir/test";
if (!is_dir($dir)) {
    mkdir($dir, 0755, recursive: true);
}
echo is_dir($dir) ? "DIR_CREATED" : "DIR_FAILED";
rmdir($dir);
"#,
    );
}

#[test]
fn test_php_glob_file_search_pattern() {
    compile_ok(
        r#"<?php
$files = glob(sys_get_temp_dir() . "/*");
echo is_array($files) ? "ARRAY" : "FALSE";
"#,
    );
}

#[test]
fn test_php_stat_filemtime_filesize_checks() {
    compile_ok(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "stat_test_");
file_put_contents($tmp, "12345");
echo filesize($tmp) . " bytes mtime=" . filemtime($tmp);
unlink($tmp);
"#,
    );
}

#[test]
fn test_php_copy_rename_file_lifecycle() {
    compile_ok(
        r#"<?php
$src = tempnam(sys_get_temp_dir(), "src_");
$dst = sys_get_temp_dir() . "/dst_file.txt";
file_put_contents($src, "copy test");
copy($src, $dst);
echo is_file($dst) ? "COPIED" : "FAIL";
unlink($src);
unlink($dst);
"#,
    );
}

#[test]
fn test_php_stream_wrapper_register_custom() {
    compile_ok(
        r#"<?php
class VariableStream {
    public static string $data = "";
    public function stream_open($path, $mode, $options, &$opened_path) { return true; }
    public function stream_write($data) { self::$data .= $data; return strlen($data); }
    public function stream_read($count) { return ""; }
}

stream_wrapper_register("var", VariableStream::class);
file_put_contents("var://buffer", "custom stream data");
echo VariableStream::$data;
stream_wrapper_unregister("var");
"#,
    );
}
