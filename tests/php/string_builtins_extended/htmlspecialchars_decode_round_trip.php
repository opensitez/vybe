<?php
// vybe-test: php/string_builtins_extended/htmlspecialchars_decode_round_trip
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$original = '<a href="url">link & text</a>';
$encoded = htmlspecialchars($original);
$decoded = htmlspecialchars_decode($encoded);
echo $decoded;
