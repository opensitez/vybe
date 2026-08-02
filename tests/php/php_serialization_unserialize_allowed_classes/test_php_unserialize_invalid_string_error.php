<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_unserialize_invalid_string_error
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

$invalidSerialized = 'a:2:{i:0;s:3:"foo";';
$restored = @unserialize($invalidSerialized);
echo $restored === false ? "UNSERIALIZE_FAILED" : "SUCCESS";
