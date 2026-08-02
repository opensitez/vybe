<?php
// vybe-test: php/references/pass_by_reference_basic
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function increment(&$val) { $val++; }
$x = 5;
increment($x);
echo $x; // 6
