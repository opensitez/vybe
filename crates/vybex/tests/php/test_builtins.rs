use super::helpers::compile_ok;

// ── String builtins ─────────────────────────────────────────
#[test] fn strlen() { compile_ok("<?php $x = strlen('hello');"); }
#[test] fn strtolower() { compile_ok("<?php $x = strtolower('HELLO');"); }
#[test] fn strtoupper() { compile_ok("<?php $x = strtoupper('hello');"); }
#[test] fn trim() { compile_ok("<?php $x = trim('  hello  ');"); }
#[test] fn ltrim() { compile_ok("<?php $x = ltrim('  hello');"); }
#[test] fn rtrim() { compile_ok("<?php $x = rtrim('hello  ');"); }
#[test] fn substr() { compile_ok("<?php $x = substr('hello', 1, 3);"); }
#[test] fn str_replace() { compile_ok("<?php $x = str_replace('o', '0', 'hello');"); }
#[test] fn explode() { compile_ok("<?php $x = explode(',', 'a,b,c');"); }
#[test] fn implode() { compile_ok("<?php $x = implode(',', ['a','b','c']);"); }
#[test] fn strpos() { compile_ok("<?php $x = strpos('hello', 'lo');"); }
#[test] fn str_contains() { compile_ok("<?php $x = str_contains('hello', 'ell');"); }
#[test] fn str_starts_with() { compile_ok("<?php $x = str_starts_with('hello', 'he');"); }
#[test] fn str_ends_with() { compile_ok("<?php $x = str_ends_with('hello', 'lo');"); }
#[test] fn str_repeat() { compile_ok("<?php $x = str_repeat('ab', 3);"); }
#[test] fn str_pad() { compile_ok("<?php $x = str_pad('42', 5, '0');"); }
#[test] fn chr_ord() { compile_ok("<?php $x = chr(65); $y = ord('A');"); }
#[test] fn ucfirst() { compile_ok("<?php $x = ucfirst('hello');"); }
#[test] fn lcfirst() { compile_ok("<?php $x = lcfirst('Hello');"); }
#[test] fn nl2br() { compile_ok("<?php $x = nl2br(\"hello\\nworld\");"); }
#[test] fn htmlspecialchars() { compile_ok("<?php $x = htmlspecialchars('<b>hi</b>');"); }
#[test] fn sprintf() { compile_ok("<?php $x = sprintf('Hello %s, age %d', 'John', 30);"); }

// ── Array builtins ──────────────────────────────────────────
#[test] fn count() { compile_ok("<?php $x = count([1,2,3]);"); }
#[test] fn array_push() { compile_ok("<?php $a = [1]; array_push($a, 2);"); }
#[test] fn array_pop() { compile_ok("<?php $a = [1,2]; $v = array_pop($a);"); }
#[test] fn array_shift() { compile_ok("<?php $a = [1,2]; $v = array_shift($a);"); }
#[test] fn array_reverse() { compile_ok("<?php $x = array_reverse([1,2,3]);"); }
#[test] fn array_slice() { compile_ok("<?php $x = array_slice([1,2,3,4], 1, 2);"); }
#[test] fn array_merge() { compile_ok("<?php $x = array_merge([1,2], [3,4]);"); }
#[test] fn array_search() { compile_ok("<?php $x = array_search(2, [1,2,3]);"); }
#[test] fn in_array() { compile_ok("<?php $x = in_array(2, [1,2,3]);"); }
#[test] fn array_keys() { compile_ok("<?php $x = array_keys(['a'=>1,'b'=>2]);"); }
#[test] fn array_values() { compile_ok("<?php $x = array_values(['a'=>1,'b'=>2]);"); }
#[test] fn sort() { compile_ok("<?php $a = [3,1,2]; sort($a);"); }
#[test] fn range() { compile_ok("<?php $x = range(1, 10);"); }
#[test] fn array_sum() { compile_ok("<?php $x = array_sum([1,2,3]);"); }
#[test] fn compact() { compile_ok("<?php $a = 1; $b = 2; $x = compact('a', 'b');"); }
#[test] fn array_key_exists() { compile_ok("<?php $x = array_key_exists('a', ['a'=>1]);"); }

// ── Callback array ops ──────────────────────────────────────
#[test] fn array_map() { compile_ok("<?php $x = array_map(fn($n) => $n * 2, [1,2,3]);"); }
#[test] fn array_filter_no_cb() { compile_ok("<?php $x = array_filter([1,0,2,null,3]);"); }
#[test] fn array_filter_cb() { compile_ok("<?php $x = array_filter([1,2,3,4], fn($n) => $n > 2);"); }
#[test] fn array_reduce() { compile_ok("<?php $x = array_reduce([1,2,3], fn($c,$i) => $c + $i, 0);"); }
#[test] fn array_walk() { compile_ok("<?php $a = [1,2,3]; array_walk($a, fn($v,$k) => $v);"); }
#[test] fn usort() { compile_ok("<?php $a = [3,1,2]; usort($a);"); }

// ── Math builtins ───────────────────────────────────────────
#[test] fn abs() { compile_ok("<?php $x = abs(-5);"); }
#[test] fn ceil() { compile_ok("<?php $x = ceil(1.2);"); }
#[test] fn floor() { compile_ok("<?php $x = floor(1.8);"); }
#[test] fn round() { compile_ok("<?php $x = round(1.5);"); }
#[test] fn sqrt() { compile_ok("<?php $x = sqrt(16);"); }
#[test] fn pow() { compile_ok("<?php $x = pow(2, 8);"); }
#[test] fn max_min() { compile_ok("<?php $a = max(1,2,3); $b = min(1,2,3);"); }
#[test] fn sin_cos_tan() { compile_ok("<?php $a = sin(1.0); $b = cos(1.0); $c = tan(1.0);"); }
#[test] fn log_exp() { compile_ok("<?php $a = log(10); $b = exp(1);"); }
#[test] fn rand() { compile_ok("<?php $x = rand();"); }

// ── Type builtins ───────────────────────────────────────────
#[test] fn intval() { compile_ok("<?php $x = intval('42');"); }
#[test] fn floatval() { compile_ok("<?php $x = floatval('3.14');"); }
#[test] fn strval() { compile_ok("<?php $x = strval(42);"); }
#[test] fn boolval() { compile_ok("<?php $x = boolval(1);"); }
#[test] fn is_null() { compile_ok("<?php $x = is_null(null);"); }
#[test] fn is_numeric() { compile_ok("<?php $x = is_numeric('42');"); }
#[test] fn is_array() { compile_ok("<?php $x = is_array([]);"); }
#[test] fn is_string() { compile_ok("<?php $x = is_string('hi');"); }
#[test] fn is_int() { compile_ok("<?php $x = is_int(42);"); }
#[test] fn is_bool() { compile_ok("<?php $x = is_bool(true);"); }
#[test] fn isset() { compile_ok("<?php $x = isset($a);"); }
#[test] fn empty() { compile_ok("<?php $x = empty($a);"); }
#[test] fn gettype() { compile_ok("<?php $x = gettype(42);"); }
#[test] fn define_defined() { compile_ok("<?php define('FOO', 42); $x = defined('FOO');"); }
#[test] fn function_exists() { compile_ok("<?php $x = function_exists('strlen');"); }
#[test] fn class_exists() { compile_ok("<?php $x = class_exists('stdClass');"); }

// ── Encoding / JSON / Crypto ────────────────────────────────
#[test] fn json_encode() { compile_ok("<?php $x = json_encode(['a'=>1]);"); }
#[test] fn json_decode() { compile_ok("<?php $x = json_decode('{\"a\":1}');"); }
#[test] fn urlencode() { compile_ok("<?php $x = urlencode('hello world');"); }
#[test] fn urldecode() { compile_ok("<?php $x = urldecode('hello%20world');"); }
#[test] fn base64_encode() { compile_ok("<?php $x = base64_encode('hello');"); }
#[test] fn base64_decode() { compile_ok("<?php $x = base64_decode('aGVsbG8=');"); }
#[test] fn md5() { compile_ok("<?php $x = md5('hello');"); }
#[test] fn sha1() { compile_ok("<?php $x = sha1('hello');"); }

// ── Regex ───────────────────────────────────────────────────
#[test] fn preg_match() { compile_ok("<?php $x = preg_match('/\\d+/', 'abc123');"); }
#[test] fn preg_replace() { compile_ok("<?php $x = preg_replace('/\\d/', 'X', 'a1b2');"); }
#[test] fn preg_split() { compile_ok("<?php $x = preg_split('/,/', 'a,b,c');"); }

// ── Filesystem / IO ─────────────────────────────────────────
#[test] fn file_exists() { compile_ok("<?php $x = file_exists('/tmp/test');"); }
#[test] fn dirname_basename() { compile_ok("<?php $x = dirname('/tmp/test.txt'); $y = basename('/tmp/test.txt');"); }
#[test] fn time() { compile_ok("<?php $t = time();"); }
#[test] fn die() { compile_ok("<?php die('goodbye');"); }
#[test] fn exit_call() { compile_ok("<?php exit(0);"); }
