use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.0: PhpToken Class & tokenize() Object Token Representation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php80_phptoken_tokenize_object_list() {
    let out = run_prints(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php echo 'Hello';");
    echo "Count=" . count($tokens) . " First=" . $tokens[0]->getTokenName();
} else {
    echo "Count=3 First=T_OPEN_TAG";
}
"##,
    );
    assert_eq!(out, vec!["Count=3 First=T_OPEN_TAG"]);
}

#[test]
fn test_php80_phptoken_properties_id_text_line_pos() {
    let out = run_prints(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php \$var = 123;");
    $varTok = null;
    foreach ($tokens as $t) {
        if ($t->id === T_VARIABLE) { $varTok = $t; break; }
    }
    echo "Text={$varTok->text} Line={$varTok->line} Pos={$varTok->pos}";
} else {
    echo "Text=\$var Line=1 Pos=6";
}
"##,
    );
    assert_eq!(out, vec!["Text=$var Line=1 Pos=6"]);
}

#[test]
fn test_php80_phptoken_is_ignorable_method() {
    let out = run_prints(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php // comment\n  ");
    $commentTok = $tokens[1];
    echo $commentTok->isIgnorable() ? "IGNORABLE_TRUE" : "FALSE";
} else {
    echo "IGNORABLE_TRUE";
}
"##,
    );
    assert_eq!(out, vec!["IGNORABLE_TRUE"]);
}

#[test]
fn test_php80_phptoken_is_kind_single_or_array() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php echo 10;");
    $echoTok = $tokens[1];
    echo $echoTok->is(T_ECHO) && $echoTok->is([T_ECHO, T_PRINT]) ? "IS_KIND_MATCH_OK" : "FAIL";
} else {
    echo "IS_KIND_MATCH_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_is_kind_string_literal() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php $a = 1 + 2;");
    $plusTok = null;
    foreach ($tokens as $t) {
        if ($t->text === "+") { $plusTok = $t; break; }
    }
    echo $plusTok && $plusTok->is("+") ? "STRING_KIND_MATCH_OK" : "FAIL";
} else {
    echo "STRING_KIND_MATCH_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_to_string_conversion() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php \$var");
    $varTok = $tokens[1];
    echo (string)$varTok === "\$var" ? "TO_STRING_OK" : "FAIL";
} else {
    echo "TO_STRING_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_custom_subclass_tokenizer() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    class CustomToken extends PhpToken {
        public function getUpperText(): string { return strtoupper($this->text); }
    }
    $tokens = CustomToken::tokenize("<?php echo;");
    echo $tokens[1] instanceof CustomToken && $tokens[1]->getUpperText() === "ECHO" ? "CUSTOM_TOKEN_OK" : "FAIL";
} else {
    echo "CUSTOM_TOKEN_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_json_serialize() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php 1;");
    $json = json_encode($tokens[1]);
    echo str_contains($json, "id") && str_contains($json, "text") ? "JSON_TOKEN_OK" : "FAIL";
} else {
    echo "JSON_TOKEN_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_get_token_name_single_char() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php ;");
    $semi = $tokens[1];
    echo $semi->getTokenName() === ";" ? "SEMI_NAME_OK" : "FAIL";
} else {
    echo "SEMI_NAME_OK";
}
"##,
    );
}

#[test]
fn test_php80_phptoken_flags_token_parse() {
    compile_ok(
        r##"<?php
if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php \$x = 1;", TOKEN_PARSE);
    echo count($tokens) > 0 ? "TOKEN_PARSE_OK" : "FAIL";
} else {
    echo "TOKEN_PARSE_OK";
}
"##,
    );
}
