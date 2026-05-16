use super::helpers::compile_ok;

// ── gzcompress / gzuncompress ─────────────────────────────────

#[test] fn gzcompress_basic() {
    compile_ok(r#"<?php
$data = "Hello, World! This is a test string for compression.";
$compressed = gzcompress($data);
echo strlen($compressed) > 0 ? 'compressed' : 'empty';
echo strlen($compressed) <= strlen($data) ? ':smaller or equal' : ':larger';
"#);
}

#[test] fn gzcompress_roundtrip() {
    compile_ok(r#"<?php
$original = str_repeat("compress me! ", 100);
$compressed = gzcompress($original);
$restored = gzuncompress($compressed);
echo $restored === $original ? 'roundtrip ok' : 'fail';
"#);
}

#[test] fn gzcompress_level() {
    compile_ok(r#"<?php
$data = str_repeat("abcdef", 200);
$fast = gzcompress($data, 1);
$best = gzcompress($data, 9);
echo strlen($fast) > 0 ? 'fast ok' : 'fail';
echo strlen($best) > 0 ? ':best ok' : ':fail';
echo strlen($best) <= strlen($fast) ? ':best <= fast' : ':unexpected';
"#);
}

#[test] fn gzuncompress_roundtrip() {
    compile_ok(r#"<?php
$texts = ["", "a", str_repeat("x", 1000), "Unicode: héllo wörld"];
foreach ($texts as $t) {
    $c = gzcompress($t);
    echo gzuncompress($c) === $t ? 'ok ' : 'fail ';
}
"#);
}

// ── gzencode / gzdecode ───────────────────────────────────────

#[test] fn gzencode_basic() {
    compile_ok(r#"<?php
$data = "Hello gzip world!";
$encoded = gzencode($data);
echo strlen($encoded) > 0 ? 'encoded' : 'empty';
// gzip has header magic bytes
echo substr($encoded, 0, 2) === "\x1f\x8b" ? ':gzip magic' : ':no magic';
"#);
}

#[test] fn gzencode_gzdecode_roundtrip() {
    compile_ok(r#"<?php
$original = str_repeat("gzip test data ", 50);
$encoded = gzencode($original);
$decoded = gzdecode($encoded);
echo $decoded === $original ? 'roundtrip ok' : 'fail';
"#);
}

#[test] fn gzencode_level() {
    compile_ok(r#"<?php
$data = str_repeat("level test ", 100);
$low  = gzencode($data, 1);
$high = gzencode($data, 9);
echo strlen($low)  > 0 ? 'low ok'  : 'fail';
echo strlen($high) > 0 ? ':high ok' : ':fail';
"#);
}

#[test] fn gzdecode_basic() {
    compile_ok(r#"<?php
$data = "decode this";
$encoded = gzencode($data);
$decoded = gzdecode($encoded);
echo $decoded;
"#);
}

// ── gzdeflate / gzinflate ─────────────────────────────────────

#[test] fn gzdeflate_basic() {
    compile_ok(r#"<?php
$data = "deflate compress test";
$deflated = gzdeflate($data);
echo strlen($deflated) > 0 ? 'deflated' : 'empty';
"#);
}

#[test] fn gzdeflate_gzinflate_roundtrip() {
    compile_ok(r#"<?php
$original = str_repeat("deflate me ", 50);
$deflated = gzdeflate($original);
$inflated = gzinflate($deflated);
echo $inflated === $original ? 'roundtrip ok' : 'fail';
"#);
}

#[test] fn gzdeflate_level() {
    compile_ok(r#"<?php
$data = str_repeat("hello world ", 100);
$d1 = gzdeflate($data, 1);
$d9 = gzdeflate($data, 9);
echo strlen($d9) <= strlen($d1) ? 'level 9 smaller' : 'unexpected';
"#);
}

#[test] fn gzinflate_basic() {
    compile_ok(r#"<?php
$data = "inflate this string";
$deflated = gzdeflate($data);
echo gzinflate($deflated);
"#);
}

// ── zlib_encode / zlib_decode ─────────────────────────────────

#[test] fn zlib_encode_decode_deflate() {
    compile_ok(r#"<?php
$data = str_repeat("zlib test ", 20);
$encoded = zlib_encode($data, ZLIB_ENCODING_DEFLATE);
$decoded = zlib_decode($encoded);
echo $decoded === $data ? 'deflate ok' : 'fail';
"#);
}

#[test] fn zlib_encode_decode_gzip() {
    compile_ok(r#"<?php
$data = "gzip via zlib_encode";
$encoded = zlib_encode($data, ZLIB_ENCODING_GZIP);
$decoded = zlib_decode($encoded);
echo $decoded === $data ? 'gzip ok' : 'fail';
"#);
}

#[test] fn zlib_encode_decode_raw() {
    compile_ok(r#"<?php
$data = str_repeat("raw deflate ", 30);
$encoded = zlib_encode($data, ZLIB_ENCODING_RAW);
$decoded = zlib_decode($encoded);
echo $decoded === $data ? 'raw ok' : 'fail';
"#);
}

// ── gzip file functions ───────────────────────────────────────

#[test] fn gzfile_write_read() {
    compile_ok(r#"<?php
$tmpfile = sys_get_temp_dir() . '/test_' . uniqid() . '.gz';
$fh = gzopen($tmpfile, 'w');
gzwrite($fh, "line one\n");
gzwrite($fh, "line two\n");
gzclose($fh);
$fh = gzopen($tmpfile, 'r');
$line1 = gzgets($fh);
$line2 = gzgets($fh);
gzclose($fh);
@unlink($tmpfile);
echo trim($line1) . ':' . trim($line2);
"#);
}

#[test] fn gzfile_compress_decompress() {
    compile_ok(r#"<?php
$tmpfile = sys_get_temp_dir() . '/test_' . uniqid() . '.gz';
$data = str_repeat("gzfile test ", 50);
$fh = gzopen($tmpfile, 'w9');
gzwrite($fh, $data);
gzclose($fh);
$fh = gzopen($tmpfile, 'r');
$restored = '';
while (!gzeof($fh)) { $restored .= gzread($fh, 1024); }
gzclose($fh);
@unlink($tmpfile);
echo $restored === $data ? 'file roundtrip ok' : 'fail';
"#);
}

// ── Compression comparison ────────────────────────────────────

#[test] fn compare_compression_ratios() {
    compile_ok(r#"<?php
$data = str_repeat("The quick brown fox jumps over the lazy dog. ", 100);
$original_size = strlen($data);
$compressed   = strlen(gzcompress($data, 9));
$deflated     = strlen(gzdeflate($data, 9));
$gzipped      = strlen(gzencode($data, 9));
echo $compressed < $original_size ? 'compress ok' : 'fail';
echo $deflated   < $original_size ? ':deflate ok' : ':fail';
echo $gzipped    < $original_size ? ':gzip ok'    : ':fail';
"#);
}

#[test] fn compression_empty_string() {
    compile_ok(r#"<?php
$empty = '';
echo gzuncompress(gzcompress($empty)) === $empty ? 'empty compress ok' : 'fail';
echo gzdecode(gzencode($empty))       === $empty ? ':empty gzip ok'    : ':fail';
echo gzinflate(gzdeflate($empty))     === $empty ? ':empty deflate ok' : ':fail';
"#);
}

#[test] fn compression_binary_data() {
    compile_ok(r#"<?php
$binary = random_bytes(256);
$compressed = gzcompress($binary);
$restored   = gzuncompress($compressed);
echo $restored === $binary ? 'binary roundtrip ok' : 'fail';
"#);
}

// ── Practical patterns ────────────────────────────────────────

#[test] fn compress_json_payload() {
    compile_ok(r#"<?php
$payload = json_encode([
    'users' => array_fill(0, 100, ['name' => 'Alice', 'email' => 'alice@example.com', 'active' => true]),
]);
$compressed = gzencode($payload, 6);
$ratio = strlen($compressed) / strlen($payload);
echo $ratio < 0.5 ? 'good ratio' : 'poor ratio';
echo gzdecode($compressed) === $payload ? ':intact' : ':corrupted';
"#);
}

#[test] fn compress_cache_value() {
    compile_ok(r#"<?php
function cacheSerialize(mixed $value): string {
    return base64_encode(gzcompress(serialize($value)));
}
function cacheUnserialize(string $cached): mixed {
    return unserialize(gzuncompress(base64_decode($cached)));
}
$original = ['data' => str_repeat('x', 500), 'count' => 42];
$cached   = cacheSerialize($original);
$restored = cacheUnserialize($cached);
echo $restored['count'] . ':' . strlen($restored['data']);
"#);
}
