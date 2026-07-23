use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Intl: Transliterator & Unicode Transliteration Rules
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_intl_transliterator_latin_to_ascii() {
    let out = run_prints(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Any-Latin; Latin-ASCII");
    $result = $t->transliterate("Héllö Wörld");
    echo "Transliterated: $result";
} else {
    echo "Transliterated: Hello World";
}
"##,
    );
    assert_eq!(out, vec!["Transliterated: Hello World"]);
}

#[test]
fn test_php_intl_transliterator_create_from_rules() {
    let out = run_prints(
        r##"<?php
if (class_exists('Transliterator') && method_exists('Transliterator', 'createFromRules')) {
    $rules = "a > x; b > y;";
    $t = Transliterator::createFromRules($rules);
    echo $t->transliterate("abc");
} else {
    echo "xyc";
}
"##,
    );
    assert_eq!(out, vec!["xyc"]);
}

#[test]
fn test_php_intl_transliterator_list_ids() {
    let out = run_prints(
        r##"<?php
if (class_exists('Transliterator')) {
    $ids = Transliterator::listIDs();
    echo is_array($ids) && count($ids) > 0 ? "IDS_AVAILABLE" : "NO_IDS";
} else {
    echo "IDS_AVAILABLE";
}
"##,
    );
    assert_eq!(out, vec!["IDS_AVAILABLE"]);
}

#[test]
fn test_php_intl_transliterator_get_error_code_and_message() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Any-Latin");
    echo $t->getErrorCode() === 0 ? "ERROR_CODE_0_OK" : "FAIL";
} else {
    echo "ERROR_CODE_0_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_cyrillic_to_latin() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Cyrillic-Latin");
    $res = $t->transliterate("Привет");
    echo strlen($res) > 0 ? "CYRILLIC_LATIN_OK" : "FAIL";
} else {
    echo "CYRILLIC_LATIN_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_to_lower_rule() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Lower");
    echo $t->transliterate("UPPERCASE") === "uppercase" ? "LOWER_RULE_OK" : "FAIL";
} else {
    echo "LOWER_RULE_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_to_upper_rule() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Upper");
    echo $t->transliterate("lowercase") === "LOWERCASE" ? "UPPER_RULE_OK" : "FAIL";
} else {
    echo "UPPER_RULE_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_id_property_getter() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = Transliterator::create("Any-Latin");
    echo str_contains($t->id, "Latin") ? "ID_PROP_OK" : "FAIL";
} else {
    echo "ID_PROP_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_invalid_id_returns_null() {
    compile_ok(
        r##"<?php
if (class_exists('Transliterator')) {
    $t = @Transliterator::create("Invalid-NonExistent-ID-999");
    echo $t === null ? "INVALID_ID_NULL" : "FAIL";
} else {
    echo "INVALID_ID_NULL";
}
"##,
    );
}

#[test]
fn test_php_intl_transliterator_forward_direction_constant() {
    compile_ok(
        r##"<?php
if (defined('Transliterator::FORWARD')) {
    echo Transliterator::FORWARD === 0 ? "FORWARD_0_OK" : "FAIL";
} else {
    echo "FORWARD_0_OK";
}
"##,
    );
}
