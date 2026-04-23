use super::helpers::compile_ok;

// ── pack / unpack basics ──────────────────────────────────────

#[test] fn pack_unsigned_char() {
    compile_ok(r#"<?php
$packed = pack('C', 65);   // 'A'
echo ord($packed) . ':' . $packed;
"#);
}

#[test] fn pack_multiple_chars() {
    compile_ok(r#"<?php
$packed = pack('CCC', 72, 101, 108);
echo $packed;  // "Hel"
"#);
}

#[test] fn pack_signed_int() {
    compile_ok(r#"<?php
$packed = pack('l', -1);
echo strlen($packed);  // 4 bytes
"#);
}

#[test] fn pack_unsigned_short() {
    compile_ok(r#"<?php
$packed = pack('n', 0x0102);  // big-endian unsigned short
echo strlen($packed) . ':' . bin2hex($packed);
"#);
}

#[test] fn pack_big_endian_long() {
    compile_ok(r#"<?php
$packed = pack('N', 16909060);  // 0x01020304 big-endian
echo bin2hex($packed);
"#);
}

#[test] fn pack_little_endian_long() {
    compile_ok(r#"<?php
$packed = pack('V', 16909060);  // 0x01020304 little-endian
echo bin2hex($packed);
"#);
}

#[test] fn pack_float() {
    compile_ok(r#"<?php
$packed = pack('f', 3.14);
echo strlen($packed);  // 4 bytes
"#);
}

#[test] fn pack_double() {
    compile_ok(r#"<?php
$packed = pack('d', 3.14159265358979);
echo strlen($packed);  // 8 bytes
"#);
}

#[test] fn pack_string() {
    compile_ok(r#"<?php
$packed = pack('A5', 'hello');
echo $packed;
$padded = pack('A10', 'hi');
echo strlen($padded);  // 10
"#);
}

#[test] fn pack_null_padded_string() {
    compile_ok(r#"<?php
$packed = pack('a5', 'hi');
echo strlen($packed) . ':' . bin2hex($packed);
"#);
}

// ── unpack ───────────────────────────────────────────────────

#[test] fn unpack_unsigned_char() {
    compile_ok(r#"<?php
$packed = pack('C', 42);
$result = unpack('Cval', $packed);
echo $result['val'];
"#);
}

#[test] fn unpack_multiple_fields() {
    compile_ok(r#"<?php
$packed = pack('CCS', 1, 2, 300);
$result = unpack('Cbyte1/Cbyte2/Sshort', $packed);
echo $result['byte1'] . ',' . $result['byte2'] . ',' . $result['short'];
"#);
}

#[test] fn unpack_big_endian_long() {
    compile_ok(r#"<?php
$packed = pack('N', 12345678);
$result = unpack('Nval', $packed);
echo $result['val'];
"#);
}

#[test] fn unpack_array_format() {
    compile_ok(r#"<?php
$packed = pack('C4', 10, 20, 30, 40);
$result = unpack('C4bytes', $packed);
echo implode(',', $result);
"#);
}

#[test] fn pack_unpack_roundtrip() {
    compile_ok(r#"<?php
$data = ['x' => 100, 'y' => 200, 'z' => 300];
$packed = pack('NNN', $data['x'], $data['y'], $data['z']);
$out = unpack('Nx/Ny/Nz', $packed);
echo $out['x'] . ',' . $out['y'] . ',' . $out['z'];
"#);
}

#[test] fn pack_unpack_float_roundtrip() {
    compile_ok(r#"<?php
$val = 3.14;
$packed = pack('d', $val);
$result = unpack('dval', $packed);
echo round($result['val'], 2);
"#);
}

// ── bin2hex / hex2bin ─────────────────────────────────────────

#[test] fn bin2hex_basic() {
    compile_ok(r#"<?php
echo bin2hex('A');       // 41
echo bin2hex('Hello');   // 48656c6c6f
echo bin2hex("\x00\xFF"); // 00ff
"#);
}

#[test] fn hex2bin_basic() {
    compile_ok(r#"<?php
echo hex2bin('48656c6c6f');  // Hello
echo hex2bin('41');          // A
"#);
}

#[test] fn bin2hex_hex2bin_roundtrip() {
    compile_ok(r#"<?php
$data = "Binary\x00data\xFF\xFE";
echo hex2bin(bin2hex($data)) === $data ? 'roundtrip ok' : 'fail';
"#);
}

#[test] fn bin2hex_random_bytes() {
    compile_ok(r#"<?php
$bytes = random_bytes(16);
$hex = bin2hex($bytes);
echo strlen($hex) === 32 ? '32 hex chars' : 'wrong length';
echo ctype_xdigit($hex) ? ':valid hex' : ':invalid hex';
"#);
}

// ── ord / chr patterns ────────────────────────────────────────

#[test] fn chr_ord_roundtrip() {
    compile_ok(r#"<?php
for ($i = 0; $i < 128; $i++) {
    if (ord(chr($i)) !== $i) { echo "fail at $i"; break; }
}
echo 'all ok';
"#);
}

#[test] fn ord_multibyte_first_byte() {
    compile_ok(r#"<?php
echo ord('A');    // 65
echo ord('a');    // 97
echo ord(' ');    // 32
echo ord("\x00"); // 0
echo ord("\xFF"); // 255
"#);
}

#[test] fn chr_control_chars() {
    compile_ok(r#"<?php
echo bin2hex(chr(0));    // 00
echo bin2hex(chr(9));    // 09 (tab)
echo bin2hex(chr(10));   // 0a (newline)
echo bin2hex(chr(13));   // 0d (carriage return)
echo bin2hex(chr(27));   // 1b (escape)
"#);
}

// ── Binary string operations ──────────────────────────────────

#[test] fn binary_string_length() {
    compile_ok(r#"<?php
$bin = "\x00\x01\x02\xFF\xFE";
echo strlen($bin);  // 5 — counts bytes not chars
"#);
}

#[test] fn binary_string_substr() {
    compile_ok(r#"<?php
$bin = pack('CCCC', 1, 2, 3, 4);
$part = substr($bin, 1, 2);
$result = unpack('C*', $part);
echo implode(',', $result);
"#);
}

#[test] fn binary_string_compare() {
    compile_ok(r#"<?php
$a = "\x01\x02\x03";
$b = "\x01\x02\x03";
$c = "\x01\x02\x04";
echo ($a === $b) ? 'equal' : 'not equal';
echo ($a === $c) ? 'equal' : ':not equal';
"#);
}

#[test] fn binary_string_search() {
    compile_ok(r#"<?php
$data = pack('CNCN', 0xAA, 0x12345678, 0xBB, 0x87654321);
$pos = strpos($data, chr(0xBB));
echo $pos > 0 ? 'found' : 'not found';
"#);
}

// ── Bit operations ────────────────────────────────────────────

#[test] fn bitwise_and() {
    compile_ok(r#"<?php
echo 0b1100 & 0b1010;  // 0b1000 = 8
echo 0xFF & 0x0F;      // 0x0F = 15
"#);
}

#[test] fn bitwise_or() {
    compile_ok(r#"<?php
echo 0b1100 | 0b1010;  // 0b1110 = 14
echo 0xF0 | 0x0F;      // 0xFF = 255
"#);
}

#[test] fn bitwise_xor() {
    compile_ok(r#"<?php
echo 0b1100 ^ 0b1010;  // 0b0110 = 6
echo 0xFF ^ 0xF0;      // 0x0F = 15
"#);
}

#[test] fn bitwise_not() {
    compile_ok(r#"<?php
echo ~0 & 0xFF;   // 255
echo ~1 & 0xFF;   // 254
"#);
}

#[test] fn bitwise_shift_left() {
    compile_ok(r#"<?php
echo 1 << 4;   // 16
echo 1 << 8;   // 256
echo 3 << 2;   // 12
"#);
}

#[test] fn bitwise_shift_right() {
    compile_ok(r#"<?php
echo 256 >> 4;  // 16
echo 128 >> 3;  // 16
echo 0xFF >> 4; // 15
"#);
}

#[test] fn bitwise_flags_pattern() {
    compile_ok(r#"<?php
const READ    = 1;
const WRITE   = 2;
const EXECUTE = 4;
$perms = READ | WRITE;
echo ($perms & READ)    ? 'can read '    : '';
echo ($perms & WRITE)   ? 'can write '   : '';
echo ($perms & EXECUTE) ? 'can execute ' : 'no execute';
"#);
}

#[test] fn bitwise_toggle_bit() {
    compile_ok(r#"<?php
$flags = 0;
$flags |= (1 << 3);  // set bit 3
echo (bool)($flags & (1 << 3)) ? 'set' : 'clear';
$flags ^= (1 << 3);  // toggle bit 3
echo (bool)($flags & (1 << 3)) ? 'set' : ':clear';
"#);
}

// ── Practical binary patterns ─────────────────────────────────

#[test] fn binary_protocol_header() {
    compile_ok(r#"<?php
// Simulate a simple binary protocol header: magic(2) + version(1) + length(4)
function buildHeader(int $version, int $length): string {
    return pack('nCN', 0xBEEF, $version, $length);
}
function parseHeader(string $data): array {
    return unpack('nmagic/Cversion/Nlength', $data);
}
$header = buildHeader(2, 1024);
$parsed = parseHeader($header);
echo dechex($parsed['magic']) . ':' . $parsed['version'] . ':' . $parsed['length'];
"#);
}

#[test] fn binary_uuid_from_bytes() {
    compile_ok(r#"<?php
function bytesToUuid(string $bytes): string {
    $hex = bin2hex($bytes);
    return vsprintf('%s%s-%s-%s-%s-%s%s%s', str_split($hex, 4));
}
$bytes = random_bytes(16);
$uuid = bytesToUuid($bytes);
echo strlen($uuid) === 36 ? 'valid uuid length' : 'wrong length';
echo substr_count($uuid, '-') === 4 ? ':four dashes' : ':wrong dashes';
"#);
}

#[test] fn binary_ip_conversion() {
    compile_ok(r#"<?php
$ip = '192.168.1.100';
$parts = explode('.', $ip);
$packed = pack('CCCC', ...$parts);
echo strlen($packed) === 4 ? '4 bytes' : 'wrong';
$unpacked = unpack('C4', $packed);
echo implode('.', $unpacked) === $ip ? ':roundtrip ok' : ':fail';
"#);
}
