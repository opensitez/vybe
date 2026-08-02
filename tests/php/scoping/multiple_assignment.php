<?php
// vybe-test: php/scoping/multiple_assignment
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$a = $b = $c = 0; echo $a + $b + $c;
