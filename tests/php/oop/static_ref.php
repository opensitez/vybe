<?php
// vybe-test: php/oop/static_ref
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class M { public static function sq($n) { return $n * $n; } } $fn = M::sq(...);
