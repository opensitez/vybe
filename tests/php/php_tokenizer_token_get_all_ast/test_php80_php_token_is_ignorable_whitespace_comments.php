<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php80_php_token_is_ignorable_whitespace_comments
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$tokens = PhpToken::tokenize("<?php // comment\n ");
$ignorableCount = 0;
foreach ($tokens as $t) {
    if ($t->isIgnorable()) $ignorableCount++;
}
echo "Ignorable tokens: $ignorableCount";
