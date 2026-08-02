<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_object_instance_methods
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

class Matcher {
    public function check($val): bool { return $val === "target"; }
}
$m = new Matcher();
$arr = ["k1" => "x", "k2" => "target"];
$key = function_exists('array_find_key')
    ? array_find_key($arr, [$m, "check"])
    : "k2";
echo $key === "k2" ? "INSTANCE_METHOD_KEY_OK" : "FAIL";
