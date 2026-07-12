use super::helpers::compile_ok;

// ── HTTP / cURL ─────────────────────────────────────────────
#[test]
fn curl_workflow() {
    compile_ok(
        r#"<?php
$ch = curl_init();
curl_setopt($ch, 'CURLOPT_URL', 'https://example.com');
$result = curl_exec($ch);
curl_close($ch);
"#,
    );
}

// ── Environment ─────────────────────────────────────────────
#[test]
fn getenv() {
    compile_ok("<?php $path = getenv('PATH');");
}
#[test]
fn phpversion() {
    compile_ok("<?php echo phpversion();");
}
#[test]
fn php_sapi_name() {
    compile_ok("<?php echo php_sapi_name();");
}
#[test]
fn getcwd() {
    compile_ok("<?php echo getcwd();");
}
#[test]
fn gethostname() {
    compile_ok("<?php echo gethostname();");
}
#[test]
fn php_uname() {
    compile_ok("<?php echo php_uname();");
}

// ── Filesystem (extended) ───────────────────────────────────
#[test]
fn copy_file() {
    compile_ok("<?php copy('src.txt', 'dst.txt');");
}
#[test]
fn rename_file() {
    compile_ok("<?php rename('old.txt', 'new.txt');");
}
#[test]
fn rmdir() {
    compile_ok("<?php rmdir('/tmp/testdir');");
}
#[test]
fn glob_dir() {
    compile_ok("<?php $files = glob('/tmp/*.txt');");
}
#[test]
fn filesize() {
    compile_ok("<?php echo filesize('test.txt');");
}
#[test]
fn tempdir() {
    compile_ok("<?php echo sys_get_temp_dir();");
}
#[test]
fn file_lines() {
    compile_ok("<?php $lines = file('test.txt');");
}
#[test]
fn pathinfo() {
    compile_ok("<?php $info = pathinfo('/tmp/test.txt');");
}

// ── File handles ────────────────────────────────────────────
#[test]
fn fopen_fwrite() {
    compile_ok(
        r#"<?php
$fp = fsockopen('localhost', 80);
fwrite($fp, "GET / HTTP/1.0\r\n\r\n");
$line = fgets($fp);
fclose($fp);
"#,
    );
}

// ── Random ──────────────────────────────────────────────────
#[test]
fn random_int() {
    compile_ok("<?php $x = random_int(1, 100);");
}
#[test]
fn random_bytes() {
    compile_ok("<?php $bytes = random_bytes(16);");
}
#[test]
fn uniqid() {
    compile_ok("<?php echo uniqid();");
}

// ── Date/Time (extended) ────────────────────────────────────
#[test]
fn usleep() {
    compile_ok("<?php usleep(1000);");
}
#[test]
fn microtime() {
    compile_ok("<?php $t = microtime();");
}

// ── Process ─────────────────────────────────────────────────
#[test]
fn exec_cmd() {
    compile_ok("<?php exec('ls -la');");
}
#[test]
fn shell_exec() {
    compile_ok("<?php $out = shell_exec('whoami');");
}

// ── XML ─────────────────────────────────────────────────────
#[test]
fn simplexml_string() {
    compile_ok("<?php $xml = simplexml_load_string('<root><item>test</item></root>');");
}
#[test]
fn simplexml_file() {
    compile_ok("<?php $xml = simplexml_load_file('data.xml');");
}

// ── UUID ────────────────────────────────────────────────────
#[test]
fn uuid_create() {
    compile_ok("<?php echo uuid_create();");
}

// ── Combined real-world ─────────────────────────────────────
#[test]
fn config_file_reader() {
    compile_ok(
        r#"<?php
$content = file_get_contents('config.json');
$config = json_decode($content);
$dbDsn = 'sqlite:' . getcwd() . '/app.db';
$pdo = new PDO($dbDsn);
"#,
    );
}

#[test]
fn file_processor() {
    compile_ok(
        r#"<?php
$files = scandir('/tmp');
foreach ($files as $file) {
    if (is_file('/tmp/' . $file)) {
        $size = filesize('/tmp/' . $file);
        $ext = pathinfo($file);
        echo $file . ': ' . $size . ' bytes';
    }
}
"#,
    );
}

#[test]
fn simple_api_client() {
    compile_ok(
        r#"<?php
$response = file_get_contents('https://api.example.com/data');
$data = json_decode($response);
echo json_encode($data);
"#,
    );
}
