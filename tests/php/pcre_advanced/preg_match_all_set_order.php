<?php
// vybe-test: php/pcre_advanced/preg_match_all_set_order
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match_all('/(\d{4})-(\d{2})/', '2024-01 and 2024-06', $m, PREG_SET_ORDER);
echo count($m);
echo $m[0][0] . ':' . $m[0][1] . ':' . $m[0][2];
echo $m[1][0] . ':' . $m[1][1] . ':' . $m[1][2];
