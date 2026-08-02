<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_strncasecmp_length_limited
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs
// vybe-test-mode: compile

echo strncasecmp("Hello World", "hello php", 5) === 0 ? "MATCH_5" : "NO_MATCH";
