<?php
// vybe-test: php/scoping/static_var
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

function counter() { $c = 0; $c++; return $c; } echo counter(); echo counter();
