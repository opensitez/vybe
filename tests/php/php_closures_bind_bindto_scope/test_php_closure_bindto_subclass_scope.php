<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_bindto_subclass_scope
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

class ParentScope {
    protected string $prot = "protected_val";
}
class ChildScope extends ParentScope {}

$fn = function() { return $this->prot; };
$child = new ChildScope();
$bound = $fn->bindTo($child, ChildScope::class);
echo $bound();
