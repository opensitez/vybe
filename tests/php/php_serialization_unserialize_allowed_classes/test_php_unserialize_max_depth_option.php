<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_unserialize_max_depth_option
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

$data = [1, [2, [3, [4]]]];
$s = serialize($data);
$restored = unserialize($s, ["max_depth" => 10]);
echo count($restored);
