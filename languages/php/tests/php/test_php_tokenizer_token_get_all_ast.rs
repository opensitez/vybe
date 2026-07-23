use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Tokenizer & Lexical Analysis — token_get_all, token_name, PhpToken::tokenize() (PHP 8.0), token IDs, line numbers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php80_php_token_tokenize_object_inspection() {
    let out = run_prints(
        r#"<?php
$tokens = PhpToken::tokenize('<?php echo "Hello";');
$names = [];
foreach ($tokens as $token) {
    if (!$token->isIgnorable()) {
        $names[] = $token->getTokenName();
    }
}
echo implode(", ", $names);
"#,
    );
    assert_eq!(
        out,
        vec!["T_OPEN_TAG, T_ECHO, T_CONSTANT_ENCAPSED_STRING, ;"]
    );
}

#[test]
fn test_php_token_get_all_legacy_array_parsing() {
    let out = run_prints(
        r#"<?php
$code = '<?php $x = 10;';
$tokens = token_get_all($code);

$tokenTypes = [];
foreach ($tokens as $tok) {
    if (is_array($tok)) {
        $tokenTypes[] = token_name($tok[0]);
    } else {
        $tokenTypes[] = $tok;
    }
}
echo implode(" ", $tokenTypes);
"#,
    );
    assert_eq!(out, vec!["T_OPEN_TAG T_VARIABLE = T_LNUMBER ;"]);
}

#[test]
fn test_php80_php_token_text_and_line_number() {
    let out = run_prints(
        r#"<?php
$tokens = PhpToken::tokenize("<?php\nclass User {}");
$classTok = null;
foreach ($tokens as $t) {
    if ($t->id === T_CLASS) {
        $classTok = $t;
        break;
    }
}
echo "Line={$classTok->line} Text={$classTok->text}";
"#,
    );
    assert_eq!(out, vec!["Line=2 Text=class"]);
}

#[test]
fn test_php80_php_token_is_given_kind() {
    compile_ok(
        r#"<?php
$tokens = PhpToken::tokenize('<?php function test() {}');
$fnTok = $tokens[1]; // T_FUNCTION or whitespace
echo $fnTok->is(T_FUNCTION) ? "IS_FUNCTION" : "NOT_FUNCTION";
"#,
    );
}

#[test]
fn test_php_token_name_resolution_valid_ids() {
    compile_ok(
        r#"<?php
echo token_name(T_VARIABLE) . " " . token_name(T_FUNCTION) . " " . token_name(T_CLASS);
"#,
    );
}

#[test]
fn test_php_token_get_all_attribute_token_recognition() {
    compile_ok(
        r#"<?php
$code = '<?php #[Attribute] class Test {}';
$tokens = token_get_all($code, TOKEN_PARSE);
$hasAttr = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_ATTRIBUTE') && $t[0] === T_ATTRIBUTE) {
        $hasAttr = true;
    }
}
echo $hasAttr ? "HAS_ATTR_TOKEN" : "LEGACY_TOKENS";
"#,
    );
}

#[test]
fn test_php_tokenizer_heredoc_nowdoc_tokens() {
    compile_ok(
        r#"<?php
$code = "<?php \$s = <<<EOT\ntext\nEOT;\n";
$tokens = token_get_all($code);
$foundHeredoc = false;
foreach ($tokens as $t) {
    if (is_array($t) && ($t[0] === T_START_HEREDOC || $t[0] === T_END_HEREDOC)) {
        $foundHeredoc = true;
    }
}
echo $foundHeredoc ? "HEREDOC_FOUND" : "NO_HEREDOC";
"#,
    );
}

#[test]
fn test_php80_php_token_is_ignorable_whitespace_comments() {
    compile_ok(
        r#"<?php
$tokens = PhpToken::tokenize("<?php // comment\n ");
$ignorableCount = 0;
foreach ($tokens as $t) {
    if ($t->isIgnorable()) $ignorableCount++;
}
echo "Ignorable tokens: $ignorableCount";
"#,
    );
}

#[test]
fn test_php_token_get_all_enum_token_php81() {
    compile_ok(
        r#"<?php
$code = '<?php enum Status { case Active; }';
$tokens = token_get_all($code);
$hasEnum = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_ENUM') && $t[0] === T_ENUM) {
        $hasEnum = true;
    }
}
echo $hasEnum ? "ENUM_TOKEN_OK" : "NO_ENUM_TOKEN";
"#,
    );
}

#[test]
fn test_php_tokenizer_match_token_php80() {
    compile_ok(
        r#"<?php
$code = '<?php $x = match($a) { 1 => "one" };';
$tokens = token_get_all($code);
$hasMatch = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_MATCH') && $t[0] === T_MATCH) {
        $hasMatch = true;
    }
}
echo $hasMatch ? "MATCH_TOKEN_OK" : "NO_MATCH_TOKEN";
"#,
    );
}
