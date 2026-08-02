<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php_token_name_resolution_valid_ids
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

echo token_name(T_VARIABLE) . " " . token_name(T_FUNCTION) . " " . token_name(T_CLASS);
