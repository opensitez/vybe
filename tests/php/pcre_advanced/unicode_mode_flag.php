<?php
// vybe-test: php/pcre_advanced/unicode_mode_flag
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$str = 'Héllo Wörld';
preg_match_all('/\p{L}+/u', $str, $m);
echo implode(' ', $m[0]);
echo preg_match('/^\p{Lu}/u', 'Ñoño') ? ':uppercase start' : ':no uppercase start';
