use super::helpers::compile_ok;

// ── Loose comparison (==) quirks ──────────────────────────────

#[test] fn loose_null_equals_false() {
    compile_ok(r#"<?php
var_dump(null == false);   // true
var_dump(null == 0);       // true
var_dump(null == "");      // true
var_dump(null == "0");     // false
var_dump(null == []);      // true
"#);
}

#[test] fn loose_zero_string_comparison() {
    compile_ok(r#"<?php
var_dump(0 == "a");    // true in PHP 7, false in PHP 8
var_dump(0 == "");     // true in PHP 7, false in PHP 8
var_dump(0 == "0");    // true
var_dump(0 == false);  // true
var_dump(0 == null);   // true
"#);
}

#[test] fn loose_string_numeric() {
    compile_ok(r#"<?php
var_dump("1" == 1);       // true
var_dump("01" == 1);      // true
var_dump("1.0" == 1);     // true
var_dump("1e2" == 100);   // true
var_dump("100" == 1e2);   // true
"#);
}

#[test] fn loose_bool_comparison() {
    compile_ok(r#"<?php
var_dump(true  == 1);      // true
var_dump(true  == "1");    // true
var_dump(true  == "any");  // true
var_dump(false == 0);      // true
var_dump(false == "");     // true
var_dump(false == "0");    // true
var_dump(false == []);     // true
var_dump(false == null);   // true
"#);
}

#[test] fn strict_vs_loose() {
    compile_ok(r#"<?php
$a = 0; $b = false;
echo ($a == $b)  ? "loose equal\n" : "loose not equal\n";
echo ($a === $b) ? "strict equal\n" : "strict not equal\n";
$c = "1"; $d = 1;
echo ($c == $d)  ? "loose equal\n" : "loose not equal\n";
echo ($c === $d) ? "strict equal\n" : "strict not equal\n";
"#);
}

#[test] fn loose_array_comparison() {
    compile_ok(r#"<?php
var_dump([] == false);   // true
var_dump([] == null);    // true
var_dump([] == 0);       // false
var_dump([0] == [false]); // true
var_dump(['a' => 1] == ['a' => true]); // true
"#);
}

// ── Type coercion in arithmetic ───────────────────────────────

#[test] fn coercion_string_to_int_arithmetic() {
    compile_ok(r#"<?php
$a = "5";
$b = $a + 3;
var_dump($b);       // int(8)
$c = "5.5" + 1;
var_dump($c);       // float(6.5)
$d = "5 apples" + 2;
var_dump($d);       // int(7)
$e = "apples" + 2;
var_dump($e);       // int(2)
"#);
}

#[test] fn coercion_bool_arithmetic() {
    compile_ok(r#"<?php
var_dump(true  + true);   // int(2)
var_dump(false + 1);      // int(1)
var_dump(true  + 0.5);    // float(1.5)
var_dump(true  * 10);     // int(10)
"#);
}

#[test] fn coercion_null_arithmetic() {
    compile_ok(r#"<?php
var_dump(null + 1);     // int(1)
var_dump(null + 1.5);   // float(1.5)
var_dump(null . "str"); // string "str"
"#);
}

#[test] fn coercion_concat_converts_to_string() {
    compile_ok(r#"<?php
$n = 42;
$f = 3.14;
$b = true;
echo $n . "," . $f . "," . $b;
echo "\n";
echo null . "null";
"#);
}

// ── settype ───────────────────────────────────────────────────

#[test] fn settype_to_int() {
    compile_ok(r#"<?php
$v = "42abc";
settype($v, 'integer');
var_dump($v);
"#);
}

#[test] fn settype_to_float() {
    compile_ok(r#"<?php
$v = "3.99";
settype($v, 'float');
var_dump($v);
"#);
}

#[test] fn settype_to_bool() {
    compile_ok(r#"<?php
$v = 0;
settype($v, 'bool');
var_dump($v);
$v2 = "hello";
settype($v2, 'boolean');
var_dump($v2);
"#);
}

#[test] fn settype_to_string() {
    compile_ok(r#"<?php
$v = 123;
settype($v, 'string');
var_dump($v);
"#);
}

#[test] fn settype_to_array() {
    compile_ok(r#"<?php
$v = 42;
settype($v, 'array');
var_dump($v);
"#);
}

// ── Casting ───────────────────────────────────────────────────

#[test] fn cast_int() {
    compile_ok(r#"<?php
var_dump((int) "42");
var_dump((int) "42abc");
var_dump((int) "abc");
var_dump((int) 3.9);
var_dump((int) true);
var_dump((int) null);
var_dump((int) false);
"#);
}

#[test] fn cast_float() {
    compile_ok(r#"<?php
var_dump((float) "3.14");
var_dump((float) "1e3");
var_dump((float) "abc");
var_dump((float) true);
var_dump((float) null);
"#);
}

#[test] fn cast_string() {
    compile_ok(r#"<?php
var_dump((string) 42);
var_dump((string) 3.14);
var_dump((string) true);
var_dump((string) false);
var_dump((string) null);
"#);
}

#[test] fn cast_bool() {
    compile_ok(r#"<?php
var_dump((bool) 1);
var_dump((bool) 0);
var_dump((bool) -1);
var_dump((bool) "");
var_dump((bool) "0");
var_dump((bool) "false");
var_dump((bool) []);
var_dump((bool) [0]);
var_dump((bool) null);
"#);
}

#[test] fn cast_array() {
    compile_ok(r#"<?php
var_dump((array) 42);
var_dump((array) "hello");
var_dump((array) null);
$obj = new stdClass(); $obj->x = 1; $obj->y = 2;
var_dump((array) $obj);
"#);
}

#[test] fn cast_object() {
    compile_ok(r#"<?php
$arr = ['name' => 'Alice', 'age' => 30];
$obj = (object) $arr;
echo $obj->name . ':' . $obj->age;
"#);
}

// ── intval / floatval / strval / boolval ──────────────────────

#[test] fn intval_bases() {
    compile_ok(r#"<?php
echo intval('0x1A', 16) . "\n";  // 26
echo intval('0b1010', 2) . "\n"; // 10
echo intval('077', 8) . "\n";    // 63
echo intval('42', 10) . "\n";    // 42
echo intval('42abc') . "\n";     // 42
"#);
}

#[test] fn floatval_variants() {
    compile_ok(r#"<?php
echo floatval('1.5e3') . "\n";   // 1500
echo floatval('  -2.5  ') . "\n"; // -2.5
echo doubleval('3.14') . "\n";
"#);
}

// ── Numeric strings ───────────────────────────────────────────

#[test] fn is_numeric_check() {
    compile_ok(r#"<?php
var_dump(is_numeric(42));
var_dump(is_numeric(3.14));
var_dump(is_numeric("42"));
var_dump(is_numeric("3.14"));
var_dump(is_numeric("1e5"));
var_dump(is_numeric("42abc"));
var_dump(is_numeric("abc"));
var_dump(is_numeric(""));
var_dump(is_numeric(null));
"#);
}

#[test] fn numeric_string_comparison() {
    compile_ok(r#"<?php
// When both strings are numeric, PHP compares numerically
var_dump("1" < "10");    // true (numeric)
var_dump("abc" < "abd"); // true (string)
var_dump("2" > "10");    // false (numeric: 2 < 10)
"#);
}

// ── Spaceship operator with types ─────────────────────────────

#[test] fn spaceship_mixed_types() {
    compile_ok(r#"<?php
echo (1 <=> 2)   . "\n"; // -1
echo (2 <=> 2)   . "\n"; // 0
echo (3 <=> 2)   . "\n"; // 1
echo ("a" <=> "b") . "\n"; // -1
echo ([1,2] <=> [1,2]) . "\n"; // 0
echo ([1,3] <=> [1,2]) . "\n"; // 1
"#);
}

// ── Type coercion in function parameters ──────────────────────

#[test] fn coercion_without_strict_types() {
    compile_ok(r#"<?php
// Without strict_types, PHP coerces args
function addNums(int $a, int $b): int { return $a + $b; }
echo addNums("3", "4");  // coerces strings to ints
echo addNums(2.9, 1.1);  // coerces floats to ints (truncates)
"#);
}

#[test] fn juggling_in_switch() {
    compile_ok(r#"<?php
$val = "0";
switch ($val) {
    case false: echo "false"; break;
    case null:  echo "null";  break;
    case 0:     echo "zero";  break;
    case "0":   echo "string zero"; break;
    default:    echo "default";
}
"#);
}
