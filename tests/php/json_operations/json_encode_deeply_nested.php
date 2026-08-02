<?php
// vybe-test: php/json_operations/json_encode_deeply_nested
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$deep = ['a' => ['b' => ['c' => ['d' => 'leaf']]]];
echo json_encode($deep);
