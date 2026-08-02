<?php
// vybe-test: php/scope_patterns/nested_closures_share_captured_ref
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$log = [];
$push = function(string $msg) use (&$log): void { $log[] = $msg; };
$pushTwice = function(string $msg) use ($push): void { $push($msg); $push($msg); };
$pushTwice('hi');
echo count($log);
