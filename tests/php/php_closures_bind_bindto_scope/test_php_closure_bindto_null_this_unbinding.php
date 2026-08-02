<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_bindto_null_this_unbinding
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

class User {
    public function getClosure() {
        return function() { return $this; };
    }
}

$u = new User();
$fn = $u->getClosure();
$unbound = $fn->bindTo(null, null);
echo is_null($unbound()) ? "UNBOUND" : "BOUND";
