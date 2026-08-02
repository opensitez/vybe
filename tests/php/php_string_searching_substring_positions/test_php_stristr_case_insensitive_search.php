<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_stristr_case_insensitive_search
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs
// vybe-test-mode: compile

$email = "USER@DOMAIN.COM";
echo stristr($email, "domain.com") ? "MATCH" : "NO_MATCH";
