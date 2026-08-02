<?php
// vybe-test: php/closures/iife
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$result = (function() { return 42; })(); echo $result;
