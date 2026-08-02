<?php
// vybe-test: php/strings/json_roundtrip
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$j = json_encode(['a'=>1]); $d = json_decode($j);
