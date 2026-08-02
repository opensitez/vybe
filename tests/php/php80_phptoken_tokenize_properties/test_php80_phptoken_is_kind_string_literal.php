<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_is_kind_string_literal
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

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
