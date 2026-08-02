<?php
// vybe-test: php/pcre_advanced/atomic_group
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

// Atomic group prevents backtracking into it
$pattern = '/(?>a|ab)c/';
echo preg_match($pattern, 'abc') ? 'matched' : 'no match';
echo preg_match($pattern, 'ac') ? 'matched' : 'no match';
