<?php
// vybe-test: php/references/pass_by_reference_string_modify
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function uppercase(&$str) { $str = strtoupper($str); }
$s = "hello";
uppercase($s);
echo $s;
