<?php
// vybe-test: php/pcre_advanced/positive_lookahead
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/\d+(?= dollars)/', 'I have 100 dollars and 50 cents', $m);
echo implode(',', $m[0]);
