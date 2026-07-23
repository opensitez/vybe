use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Stream Filters & Transformations — stream_filter_append, stream_filter_prepend, stream_filter_register, string.rot13, convert.base64-encode
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_filter_append_rot13() {
    let out = run_prints(
        r#"<?php
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "string.rot13");

fwrite($stream, "Hello World!");
rewind($stream);

$filtered = stream_get_contents($stream);
fclose($stream);

echo $filtered;
"#,
    );
    assert_eq!(out, vec!["Uryyb Jbeyq!"]);
}

#[test]
fn test_php_stream_filter_append_base64_encode() {
    let out = run_prints(
        r#"<?php
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "convert.base64-encode");

fwrite($stream, "PHP Stream Filter");
rewind($stream);

$encoded = stream_get_contents($stream);
fclose($stream);

echo trim($encoded);
"#,
    );
    assert_eq!(out, vec!["UEhQIFN0cmVhbSBGaWx0ZXI="]);
}

#[test]
fn test_php_stream_filter_toupper_transformation() {
    let out = run_prints(
        r#"<?php
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "string.toupper");

fwrite($stream, "lowercase text");
rewind($stream);

$upper = stream_get_contents($stream);
fclose($stream);

echo $upper;
"#,
    );
    assert_eq!(out, vec!["LOWERCASE TEXT"]);
}

#[test]
fn test_php_stream_get_filters_list() {
    compile_ok(
        r#"<?php
$filters = stream_get_filters();
echo in_array("string.rot13", $filters) ? "ROT13_FILTER_AVAILABLE" : "NO_ROT13";
"#,
    );
}

#[test]
fn test_php_custom_stream_filter_user_filter() {
    compile_ok(
        r#"<?php
class StripVowelsFilter extends php_user_filter {
    public function filter($in, $out, &$consumed, $closing): int {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = preg_replace('/[aeiouAEIOU]/', '', $bucket->data);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }
        return PSFS_PASS_ON;
    }
}

stream_filter_register("strip_vowels", StripVowelsFilter::class);
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "strip_vowels");

fwrite($stream, "Hello World");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
"#,
    );
}

#[test]
fn test_php_stream_filter_remove_resource() {
    compile_ok(
        r#"<?php
$stream = fopen("php://memory", "r+");
$filter = stream_filter_append($stream, "string.rot13");
stream_filter_remove($filter);
fwrite($stream, "Normal text");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
"#,
    );
}

#[test]
fn test_php_stream_filter_prepend_order() {
    compile_ok(
        r#"<?php
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "string.rot13");
stream_filter_prepend($stream, "string.toupper");
fwrite($stream, "abc");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
"#,
    );
}

#[test]
fn test_php_stream_filter_base64_decode() {
    compile_ok(
        r#"<?php
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "convert.base64-decode");
fwrite($stream, "UEhQ");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
"#,
    );
}

#[test]
fn test_php_stream_copy_to_stream_with_filter() {
    compile_ok(
        r#"<?php
$src = fopen("php://memory", "r+");
$dst = fopen("php://memory", "r+");
fwrite($src, "Transfer Data");
rewind($src);

stream_filter_append($dst, "string.toupper");
stream_copy_to_stream($src, $dst);
rewind($dst);

echo stream_get_contents($dst);
fclose($src);
fclose($dst);
"#,
    );
}

#[test]
fn test_php_stream_bucket_new_creation() {
    compile_ok(
        r#"<?php
$stream = fopen("php://memory", "r+");
$bucket = stream_bucket_new($stream, "Bucket content");
echo is_object($bucket) ? "BUCKET_CREATED" : "FAIL";
fclose($stream);
"#,
    );
}
