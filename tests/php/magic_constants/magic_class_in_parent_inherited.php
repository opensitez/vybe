<?php
// vybe-test: php/magic_constants/magic_class_in_parent_inherited
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Base {
    public function whoAmI(): string { return __CLASS__; }
}
class Child extends Base {}
$c = new Child();
echo $c->whoAmI(); // "Base" — __CLASS__ resolves at definition time
