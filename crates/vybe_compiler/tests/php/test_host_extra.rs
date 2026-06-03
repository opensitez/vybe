use super::helpers::{compile_ok, run_prints};

// ── Convert / Math extras ───────────────────────────────────
#[test]
fn dechex() {
    compile_ok("<?php echo dechex(255);");
}
#[test]
fn decoct() {
    compile_ok("<?php echo decoct(8);");
}
#[test]
fn is_finite() {
    compile_ok("<?php echo is_finite(1.0);");
}
#[test]
fn is_nan() {
    compile_ok("<?php echo is_nan(NAN);");
}
#[test]
fn hypot() {
    compile_ok("<?php echo hypot(3, 4);");
}
#[test]
fn intdiv() {
    compile_ok("<?php echo intdiv(7, 2);");
}
#[test]
fn fmod() {
    compile_ok("<?php echo fmod(10.5, 3.2);");
}
#[test]
fn fdiv() {
    compile_ok("<?php echo fdiv(10, 3);");
}
#[test]
fn pi_const() {
    compile_ok("<?php echo pi();");
}

// ── Network ─────────────────────────────────────────────────
#[test]
fn gethostbyname() {
    compile_ok("<?php $ip = gethostbyname('localhost');");
}

// ── CLI / IO ────────────────────────────────────────────────
#[test]
fn readline() {
    compile_ok("<?php $input = readline();");
}
#[test]
fn error_log() {
    compile_ok("<?php error_log('something went wrong');");
}
#[test]
fn trigger_error() {
    compile_ok("<?php trigger_error('warning', E_USER_WARNING);");
}
#[test]
fn phpinfo() {
    compile_ok("<?php phpinfo();");
}
#[test]
fn get_current_user() {
    compile_ok("<?php echo get_current_user();");
}

// ── Filesystem extras ───────────────────────────────────────
#[test]
fn stat_file() {
    compile_ok("<?php $info = stat('/tmp/test.txt');");
}
#[test]
fn readdir() {
    compile_ok("<?php $entries = readdir('/tmp');");
}
#[test]
fn filesystem_stat_helpers_runtime() {
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let root = std::env::temp_dir().join(format!("vybex_php_fs_{unique}"));
    fs::create_dir_all(&root).unwrap();
    let file = root.join("sample.txt");
    fs::write(&file, "hello").unwrap();

    let file_path = file
        .to_string_lossy()
        .replace('\\', "\\\\")
        .replace('\'', "\\'");
    let dir_path = root
        .to_string_lossy()
        .replace('\\', "\\\\")
        .replace('\'', "\\'");
    let out = run_prints(&format!(
        r#"<?php
$file = '{file_path}';
$dir = '{dir_path}';
echo file_exists($file) ? 't' : 'f';
echo is_file($file) ? 't' : 'f';
echo is_dir($dir) ? 't' : 'f';
echo filesize($file);
echo filemtime($file) > 0 ? 't' : 'f';
"#
    ));

    assert_eq!(out, vec!["t", "t", "t", "5", "t"]);

    let _ = fs::remove_dir_all(&root);
}

#[test]
fn directory_iterator_read_and_close() {
    let out = run_prints(
        r#"<?php
$dir = dir('.');
$entry = $dir->read();
if ($entry !== false) echo 'ok';
$dir->close();
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn glob_runtime_matches_suffix_in_real_directory() {
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let root = std::env::temp_dir().join(format!("vybex_php_glob_{unique}"));
    fs::create_dir_all(&root).unwrap();
    fs::write(root.join("one.knt"), "a").unwrap();
    fs::write(root.join("two.knt"), "b").unwrap();
    fs::write(root.join("skip.txt"), "c").unwrap();

    let glob_path = root
        .join("*.knt")
        .to_string_lossy()
        .replace('\\', "\\\\")
        .replace('\'', "\\'");
    let out = run_prints(&format!(
        r#"<?php
$files = glob('{glob_path}');
echo count($files);
foreach ($files as $file) echo basename($file);
"#
    ));

    assert_eq!(out.first().map(String::as_str), Some("2"));
    let mut names = out.into_iter().skip(1).collect::<Vec<_>>();
    names.sort();
    assert_eq!(names, vec!["one.knt".to_string(), "two.knt".to_string()]);

    let _ = fs::remove_dir_all(&root);
}

#[test]
fn define_and_defined_runtime() {
    let out = run_prints(
        r#"<?php
echo defined('SIZESTEP') ? 'f0' : 't0';
define('SIZESTEP', 1024.0);
echo defined('SIZESTEP') ? 't1' : 'f1';
echo SIZESTEP;
"#,
    );
    assert_eq!(out, vec!["t0", "t1", "1024"]);
}

#[test]
fn mixed_case_php_builtin_lookup_runtime() {
    let out = run_prints(
        r#"<?php
echo urlEncode('a b') === urlencode('a b') ? 'ok' : 'bad';
echo rawUrlEncode('c d') === rawurlencode('c d') ? 'ok' : 'bad';
"#,
    );
    assert_eq!(out, vec!["ok", "ok"]);
}

#[test]
fn url_decode_variants_runtime() {
    let out = run_prints(
        r#"<?php
echo urldecode('a+b');
echo rawurldecode('a+b');
"#,
    );
    assert_eq!(out, vec!["a b", "a+b"]);
}

#[test]
fn symlink_helpers_and_pathinfo_runtime() {
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    #[cfg(unix)]
    {
        use std::os::unix::fs::symlink;

        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("vybex_php_symlink_{unique}"));
        fs::create_dir_all(&root).unwrap();
        let target = root.join("target.txt");
        fs::write(&target, "ok").unwrap();
        let link = root.join("link.txt");
        symlink(&target, &link).unwrap();

        let link_path = link
            .to_string_lossy()
            .replace('\\', "\\\\")
            .replace('\'', "\\'");
        let target_path = target
            .to_string_lossy()
            .replace('\\', "\\\\")
            .replace('"', "\\\"");
        let out = run_prints(&format!(
            r#"<?php
    $link = '{link_path}';
echo is_link($link) ? 't' : 'f';
echo readlink($link);
$info = pathinfo(readlink($link));
    echo $info['dirname'];
    echo $info['basename'];
    echo $info['filename'];
    echo $info['extension'];
"#
        ));

        assert_eq!(
            out,
            vec![
                "t".to_string(),
                target_path,
                root.to_string_lossy().to_string(),
                "target.txt".to_string(),
                "target".to_string(),
                "txt".to_string(),
            ]
        );

        let _ = fs::remove_dir_all(&root);
    }
}

// ── HTTP extended ───────────────────────────────────────────
#[test]
fn fetch_api() {
    compile_ok("<?php $resp = fetch('https://api.example.com/data');");
}

// ── DateTime ────────────────────────────────────────────────
#[test]
fn datetime_now() {
    compile_ok("<?php $now = new DateTime(); echo $now->format('Y-m-d');");
}
#[test]
fn datetime_parse() {
    compile_ok("<?php $dt = new DateTime('2024-01-15'); echo $dt->format('Y-m-d');");
}
#[test]
fn datetime_immutable() {
    compile_ok("<?php $dt = new DateTimeImmutable();");
}
#[test]
fn datetime_timestamp() {
    compile_ok("<?php $dt = new DateTime(); $ts = $dt->getTimestamp();");
}

// ── SplStack ────────────────────────────────────────────────
#[test]
fn spl_stack() {
    compile_ok(
        r#"<?php
$stack = new SplStack();
$stack->push('a');
$stack->push('b');
$stack->push('c');
$top = $stack->pop();
echo $top;
"#,
    );
}

// ── SplQueue ────────────────────────────────────────────────
#[test]
fn spl_queue() {
    compile_ok(
        r#"<?php
$queue = new SplQueue();
$queue->enqueue('first');
$queue->enqueue('second');
$item = $queue->dequeue();
echo $item;
$next = $queue->peek();
"#,
    );
}

// ── Real-world combined ─────────────────────────────────────
#[test]
fn datetime_db_pattern() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:app.db');
$now = new DateTime();
$timestamp = $now->format('Y-m-d H:i:s');
$pdo->exec("INSERT INTO logs (created_at) VALUES ('" . $timestamp . "')");
"#,
    );
}

#[test]
fn cli_app() {
    compile_ok(
        r#"<?php
echo "Running on: " . php_uname() . "\n";
echo "User: " . get_current_user() . "\n";
echo "CWD: " . getcwd() . "\n";
echo "PHP: " . phpversion() . "\n";
$name = readline();
echo "Hello, " . $name . "!\n";
"#,
    );
}

#[test]
fn hex_color() {
    compile_ok(
        r#"<?php
function hexColor($r, $g, $b) {
    return '#' . str_pad(dechex($r), 2, '0') . str_pad(dechex($g), 2, '0') . str_pad(dechex($b), 2, '0');
}
echo hexColor(255, 128, 0);
"#,
    );
}

#[test]
fn stack_based_calculator() {
    compile_ok(
        r#"<?php
$stack = new SplStack();
$tokens = explode(' ', '3 4 + 2 *');
foreach ($tokens as $token) {
    if (is_numeric($token)) {
        $stack->push(intval($token));
    } else {
        $b = $stack->pop();
        $a = $stack->pop();
        $result = match($token) {
            '+' => $a + $b,
            '-' => $a - $b,
            '*' => $a * $b,
            '/' => $a / $b,
            default => 0
        };
        $stack->push($result);
    }
}
echo $stack->pop();
"#,
    );
}
