<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_url_rewriter_tags_ini
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

$tags = ini_get("url_rewriter.tags");
echo is_string($tags) ? "URL_REWRITER_TAGS_INI_OK" : "FAIL";
