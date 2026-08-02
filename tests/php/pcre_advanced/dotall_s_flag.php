<?php
// vybe-test: php/pcre_advanced/dotall_s_flag
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$text = "start\nmiddle\nend";
echo preg_match('/start.+end/', $text) ? 'matched' : 'no match';
echo preg_match('/start.+end/s', $text) ? 'matched' : 'no match';
