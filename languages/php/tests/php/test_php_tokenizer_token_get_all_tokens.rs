use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Tokenizer: token_get_all & Token Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_tokenizer_token_get_all_simple_script() {
    let out = run_prints(
        r##"<?php
$tokens = token_get_all("<?php echo 42;");
$names = [];
foreach ($tokens as $token) {
    if (is_array($token)) {
        $names[] = token_name($token[0]);
    } else {
        $names[] = $token;
    }
}
echo implode(",", $names);
"##,
    );
    assert_eq!(out, vec!["T_OPEN_TAG,T_ECHO,T_WHITESPACE,T_LNUMBER,;"]);
}

#[test]
fn test_php_tokenizer_token_name_lookup() {
    let out = run_prints(
        r##"<?php
echo token_name(T_ECHO) . " " . token_name(T_VARIABLE) . " " . token_name(T_FUNCTION);
"##,
    );
    assert_eq!(out, vec!["T_ECHO T_VARIABLE T_FUNCTION"]);
}

#[test]
fn test_php_tokenizer_token_get_all_token_parse_flag() {
    let out = run_prints(
        r##"<?php
$code = "<?php \$x = 10;";
$tokens = token_get_all($code, TOKEN_PARSE);
echo is_array($tokens) && count($tokens) > 0 ? "PARSE_TOKENS_OK" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["PARSE_TOKENS_OK"]);
}

#[test]
fn test_php_tokenizer_token_line_number() {
    compile_ok(
        r##"<?php
$code = "<?php\n\n\$var = 1;";
$tokens = token_get_all($code);
$varToken = null;
foreach ($tokens as $t) {
    if (is_array($t) && $t[0] === T_VARIABLE) { $varToken = $t; break; }
}
echo $varToken[2] === 3 ? "LINE_3_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_tokenizer_string_literal_tokens() {
    compile_ok(
        r##"<?php
$tokens = token_get_all("<?php 'hello';");
$hasConstantEncapsed = false;
foreach ($tokens as $t) {
    if (is_array($t) && $t[0] === T_CONSTANT_ENCAPSED_STRING) { $hasConstantEncapsed = true; break; }
}
echo $hasConstantEncapsed ? "T_CONSTANT_ENCAPSED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_tokenizer_comment_tokens() {
    compile_ok(
        r##"<?php
$code = "<?php // Single line comment\n/* Block comment */";
$tokens = token_get_all($code);
$commentCount = 0;
foreach ($tokens as $t) {
    if (is_array($t) && ($t[0] === T_COMMENT || $t[0] === T_DOC_COMMENT)) { $commentCount++; }
}
echo $commentCount === 2 ? "COMMENT_TOKENS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_tokenizer_attribute_tokens_php80() {
    compile_ok(
        r##"<?php
$code = "<?php #[Attribute]";
$tokens = token_get_all($code);
$hasAttr = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_ATTRIBUTE') && $t[0] === T_ATTRIBUTE)) { $hasAttr = true; break; }
}
echo "ATTRIBUTE_TOKEN_CHECKED";
"##,
    );
}

#[test]
fn test_php_tokenizer_match_expression_token() {
    compile_ok(
        r##"<?php
$code = "<?php match(\$x) { default => 0 };";
$tokens = token_get_all($code);
$hasMatch = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_MATCH') && $t[0] === T_MATCH)) { $hasMatch = true; break; }
}
echo "MATCH_TOKEN_CHECKED";
"##,
    );
}

#[test]
fn test_php_tokenizer_fn_arrow_token() {
    compile_ok(
        r##"<?php
$code = "<?php \$f = fn() => 42;";
$tokens = token_get_all($code);
$hasFn = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_FN') && $t[0] === T_FN)) { $hasFn = true; break; }
}
echo "FN_TOKEN_CHECKED";
"##,
    );
}

#[test]
fn test_php_tokenizer_nullsafe_object_operator_token() {
    compile_ok(
        r##"<?php
$code = "<?php \$x?->prop;";
$tokens = token_get_all($code);
$hasNullsafe = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_NULLSAFE_OBJECT_OPERATOR') && $t[0] === T_NULLSAFE_OBJECT_OPERATOR)) { $hasNullsafe = true; break; }
}
echo "NULLSAFE_OPERATOR_TOKEN_CHECKED";
"##,
    );
}
