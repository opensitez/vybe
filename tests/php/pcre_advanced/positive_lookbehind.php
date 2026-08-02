<?php
// vybe-test: php/pcre_advanced/positive_lookbehind
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/(?<=USD )\d+/', 'USD 100 EUR 200 USD 300', $m);
echo implode(',', $m[0]);
