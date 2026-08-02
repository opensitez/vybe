<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php_tokenizer_heredoc_nowdoc_tokens
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$code = "<?php \$s = <<<EOT\ntext\nEOT;\n";
$tokens = token_get_all($code);
$foundHeredoc = false;
foreach ($tokens as $t) {
    if (is_array($t) && ($t[0] === T_START_HEREDOC || $t[0] === T_END_HEREDOC)) {
        $foundHeredoc = true;
    }
}
echo $foundHeredoc ? "HEREDOC_FOUND" : "NO_HEREDOC";
