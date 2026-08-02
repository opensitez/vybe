<?php
// vybe-test: php/magic_constants/magic_method_vs_function
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Util {
    public function run(): string { return __METHOD__; }
}
function standalone(): string { return __FUNCTION__; }
echo (new Util())->run();
echo ':';
echo standalone();
