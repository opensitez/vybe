<?php
// vybe-test: php/closures_advanced/closure_call_method
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Secret { private string $value = 'hidden'; }
$fn = function() { return $this->value; };
echo $fn->call(new Secret());
