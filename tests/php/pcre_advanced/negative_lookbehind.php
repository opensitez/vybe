<?php
// vybe-test: php/pcre_advanced/negative_lookbehind
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/(?<!USD )\b\d+/', 'USD 100 EUR 200 CAD 300', $m);
echo implode(',', $m[0]);
