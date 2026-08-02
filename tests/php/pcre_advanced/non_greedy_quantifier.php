<?php
// vybe-test: php/pcre_advanced/non_greedy_quantifier
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$html = '<b>bold</b> and <b>more</b>';
preg_match_all('/<b>.*?<\/b>/', $html, $m);
echo count($m[0]);
echo implode(',', $m[0]);
