<?php
// vybe-test: php/pcre_advanced/non_capturing_group
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

preg_match('/(?:foo|bar)(baz)/', 'foobaz', $m);
echo $m[0];
echo $m[1];
echo isset($m[2]) ? 'has 2' : 'no 2';
