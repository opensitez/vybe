<?php
// vybe-test: php/references/closure_use_by_reference
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$counter = 0;
$inc = function() use (&$counter) { $counter++; };
$inc(); $inc(); $inc();
echo $counter;
