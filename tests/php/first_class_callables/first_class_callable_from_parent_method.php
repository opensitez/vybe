<?php
// vybe-test: php/first_class_callables/first_class_callable_from_parent_method
// origin: languages/php/tests/php/test_first_class_callables.rs
// vybe-test-mode: compile

class Base {
    public function transform(string $s): string { return strtoupper($s); }
}
class Child extends Base {
    public function getParentTransform(): callable {
        return parent::transform(...);
    }
}
$child = new Child();
$fn = $child->getParentTransform();
echo $fn('hello');
