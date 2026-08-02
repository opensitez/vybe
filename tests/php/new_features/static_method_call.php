<?php
// vybe-test: php/new_features/static_method_call
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

class M { public static function sq($n) { return $n * $n; } } echo M::sq(5);
