<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_hosts_ini
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

$hosts = ini_get("url_rewriter.hosts");
echo is_string($hosts) || $hosts === false ? "URL_REWRITER_HOSTS_INI_OK" : "FAIL";
