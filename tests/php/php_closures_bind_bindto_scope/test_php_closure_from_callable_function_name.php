<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_from_callable_function_name
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

$c = Closure::fromCallable("strlen");
echo $c("test");
