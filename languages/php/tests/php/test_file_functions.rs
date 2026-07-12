use super::helpers::compile_ok;

// ── file_get_contents ────────────────────────────────────────────

#[test]
fn file_get_contents_basic() {
    compile_ok(
        r#"<?php
$contents = file_get_contents('/etc/hostname');
echo is_string($contents) ? 'string' : 'not string';
"#,
    );
}

#[test]
fn file_get_contents_with_context() {
    compile_ok(
        r#"<?php
$opts = ['http' => ['method' => 'GET', 'header' => 'Accept: text/html']];
$ctx = stream_context_create($opts);
$result = file_get_contents('http://example.com', false, $ctx);
echo is_string($result) || $result === false ? 'ok' : 'fail';
"#,
    );
}

// ── file_put_contents ────────────────────────────────────────────

#[test]
fn file_put_contents_basic() {
    compile_ok(
        r#"<?php
$bytes = file_put_contents('/tmp/test_vybe.txt', 'hello world');
echo $bytes !== false ? 'wrote' : 'failed';
"#,
    );
}

// ── file ─────────────────────────────────────────────────────────

#[test]
fn file_read_lines() {
    compile_ok(
        r#"<?php
$lines = file('/etc/hostname', FILE_IGNORE_NEW_LINES);
echo is_array($lines) ? 'array' : 'not array';
"#,
    );
}

// ── file_exists / is_file / is_dir ───────────────────────────────

#[test]
fn file_exists_check() {
    compile_ok(
        r#"<?php
$exists = file_exists('/etc/hostname');
echo is_bool($exists) ? 'bool' : 'not bool';
"#,
    );
}

#[test]
fn is_file_check() {
    compile_ok(
        r#"<?php
echo is_file('/etc/hostname') ? 'file' : 'not file';
echo is_file('/etc') ? 'file' : 'not file';
"#,
    );
}

#[test]
fn is_dir_check() {
    compile_ok(
        r#"<?php
echo is_dir('/etc') ? 'dir' : 'not dir';
echo is_dir('/etc/hostname') ? 'dir' : 'not dir';
"#,
    );
}

// ── is_readable / is_writable ────────────────────────────────────

#[test]
fn is_readable_check() {
    compile_ok(
        r#"<?php
$r = is_readable('/etc/hostname');
echo is_bool($r) ? 'bool' : 'not bool';
"#,
    );
}

#[test]
fn is_writable_check() {
    compile_ok(
        r#"<?php
$w = is_writable('/tmp');
echo is_bool($w) ? 'bool' : 'not bool';
"#,
    );
}

// ── dirname / basename / pathinfo ────────────────────────────────

#[test]
fn dirname_basic() {
    compile_ok(
        r#"<?php
$dir = dirname('/var/www/html/index.php');
echo $dir;
$nested = dirname('/a/b/c/d.txt', 2);
echo $nested;
"#,
    );
}

#[test]
fn basename_basic() {
    compile_ok(
        r#"<?php
echo basename('/var/www/html/index.php');
echo basename('/var/www/html/index.php', '.php');
"#,
    );
}

#[test]
fn pathinfo_components() {
    compile_ok(
        r#"<?php
$info = pathinfo('/var/www/html/index.php');
echo $info['dirname'];
echo $info['basename'];
echo $info['extension'];
echo $info['filename'];
echo pathinfo('/var/www/index.php', PATHINFO_EXTENSION);
"#,
    );
}

// ── realpath ─────────────────────────────────────────────────────

#[test]
fn realpath_basic() {
    compile_ok(
        r#"<?php
$real = realpath('/etc/../etc/hostname');
echo is_string($real) || $real === false ? 'ok' : 'fail';
"#,
    );
}

// ── glob ─────────────────────────────────────────────────────────

#[test]
fn glob_pattern() {
    compile_ok(
        r#"<?php
$files = glob('/tmp/*.txt');
echo is_array($files) ? 'array' : 'not array';
"#,
    );
}

// ── scandir ──────────────────────────────────────────────────────

#[test]
fn scandir_directory() {
    compile_ok(
        r#"<?php
$entries = scandir('/tmp');
echo is_array($entries) ? 'array' : 'not array';
echo in_array('.', $entries) ? 'has dot' : 'no dot';
"#,
    );
}

// ── mkdir / rmdir ─────────────────────────────────────────────────

#[test]
fn mkdir_and_rmdir() {
    compile_ok(
        r#"<?php
$path = '/tmp/vybe_test_dir_' . getmypid();
if (!is_dir($path)) {
    $ok = mkdir($path, 0755);
    echo $ok ? 'created' : 'failed';
    $ok2 = rmdir($path);
    echo $ok2 ? ':removed' : ':rmdir failed';
}
"#,
    );
}

// ── unlink ────────────────────────────────────────────────────────

#[test]
fn unlink_file() {
    compile_ok(
        r#"<?php
$path = '/tmp/vybe_unlink_' . getmypid() . '.txt';
file_put_contents($path, 'tmp');
$ok = unlink($path);
echo $ok ? 'deleted' : 'failed';
"#,
    );
}

// ── rename ────────────────────────────────────────────────────────

#[test]
fn rename_file() {
    compile_ok(
        r#"<?php
$src = '/tmp/vybe_rename_src_' . getmypid() . '.txt';
$dst = '/tmp/vybe_rename_dst_' . getmypid() . '.txt';
file_put_contents($src, 'data');
$ok = rename($src, $dst);
echo $ok ? 'renamed' : 'failed';
if (file_exists($dst)) unlink($dst);
"#,
    );
}

// ── copy ─────────────────────────────────────────────────────────

#[test]
fn copy_file() {
    compile_ok(
        r#"<?php
$src = '/tmp/vybe_copy_src_' . getmypid() . '.txt';
$dst = '/tmp/vybe_copy_dst_' . getmypid() . '.txt';
file_put_contents($src, 'data');
$ok = copy($src, $dst);
echo $ok ? 'copied' : 'failed';
if (file_exists($src)) unlink($src);
if (file_exists($dst)) unlink($dst);
"#,
    );
}
