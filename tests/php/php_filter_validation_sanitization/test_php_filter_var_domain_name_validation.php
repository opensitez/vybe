<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_domain_name_validation
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$domain = "subdomain.example.co.uk";
echo filter_var($domain, FILTER_VALIDATE_DOMAIN) ? "DOMAIN_OK" : "DOMAIN_FAIL";
