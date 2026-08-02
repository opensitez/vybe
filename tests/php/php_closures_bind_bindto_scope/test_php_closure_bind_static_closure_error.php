<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_bind_static_closure_error
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

$staticFn = static function() { return "static"; };
$bound = @$staticFn->bindTo(new stdClass());
