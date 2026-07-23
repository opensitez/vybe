use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Type Juggling, Coercion & Strictness — Type casting, boolean truthiness, strict comparisons, type conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_string_to_numeric_coercion() {
    let out = run_prints(
        r#"<?php
$strInt = "42";
$strFloat = "3.14";
$strMixed = "100apple";

echo ((int)$strInt + 10) . " | " . ((float)$strFloat * 2) . " | " . (int)$strMixed;
"#,
    );
    assert_eq!(out, vec!["52 | 6.28 | 100"]);
}

#[test]
fn test_php_boolean_truthiness_falsiness() {
    let out = run_prints(
        r#"<?php
$falsyValues = [0, 0.0, "", "0", [], null, false];
$falsyCount = 0;
foreach ($falsyValues as $v) {
    if (!$v) $falsyCount++;
}
echo "Falsy Count: $falsyCount / " . count($falsyValues);
"#,
    );
    assert_eq!(out, vec!["Falsy Count: 7 / 7"]);
}

#[test]
fn test_php_explicit_type_casting_syntax() {
    let out = run_prints(
        r#"<?php
$val = "123.45";
$asInt = (int)$val;
$asFloat = (float)$val;
$asBool = (bool)$val;
$asArr = (array)$val;
$asObj = (object)$val;

echo "$asInt | $asFloat | " . ($asBool ? "1" : "0") . " | " . $asArr[0] . " | " . $asObj->scalar;
"#,
    );
    assert_eq!(out, vec!["123 | 123.45 | 1 | 123.45 | 123.45"]);
}

#[test]
fn test_php_strict_equality_vs_loose_equality() {
    let out = run_prints(
        r#"<?php
$a = 0;
$b = "0";
$c = "0.0";

echo ($a == $b ? "1" : "0");
echo ($a === $b ? "1" : "0");
echo ($b == $c ? "1" : "0");
echo ($b === $c ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["1010"]);
}

#[test]
fn test_php_settype_in_place_conversion() {
    compile_ok(
        r#"<?php
$foo = "5bar";
settype($foo, "integer");
echo $foo === 5 ? "SETTYPE_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_gettype_builtin_returns() {
    compile_ok(
        r#"<?php
echo gettype(1) . " " . gettype(1.0) . " " . gettype("a") . " " . gettype([]) . " " . gettype(null);
"#,
    );
}

#[test]
fn test_php_is_scalar_primitive_check() {
    compile_ok(
        r#"<?php
echo is_scalar("str") && is_scalar(123) && is_scalar(true) && !is_scalar([]) ? "SCALAR_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_float_comparison_epsilon_tolerance() {
    compile_ok(
        r#"<?php
$a = 0.1 + 0.2;
$b = 0.3;
$epsilon = 0.00001;
echo abs($a - $b) < $epsilon ? "FLOAT_EQUAL" : "FLOAT_NOT_EQUAL";
"#,
    );
}

#[test]
fn test_php_array_cast_from_object() {
    compile_ok(
        r#"<?php
class User { public string $name = "Alice"; }
$arr = (array)(new User());
echo $arr["name"];
"#,
    );
}

#[test]
fn test_php_intval_floatval_strval_boolval() {
    compile_ok(
        r#"<?php
echo intval("99") + floatval("0.5") + strlen(strval(100)) + (boolval("true") ? 1 : 0);
"#,
    );
}
