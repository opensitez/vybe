<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_serialize_unserialize
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

$fixed = new SplFixedArray(2);
$fixed[0] = "val_a";
$s = serialize($fixed);
$restored = unserialize($s);
echo $restored[0] === "val_a" ? "SERIALIZE_FIXED_OK" : "FAIL";
