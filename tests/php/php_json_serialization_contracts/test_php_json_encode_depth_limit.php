<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_encode_depth_limit
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs
// vybe-test-mode: compile

$nested = [[[["deep"]]]];
$json = json_encode($nested, depth: 2);
if ($json === false && json_last_error() === JSON_ERROR_DEPTH) {
    echo "Exceeded maximum depth";
}
