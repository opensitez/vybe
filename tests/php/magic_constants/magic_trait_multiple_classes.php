<?php
// vybe-test: php/magic_constants/magic_trait_multiple_classes
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

trait Logging {
    public function source(): string { return __TRAIT__; }
}
class A { use Logging; }
class B { use Logging; }
echo (new A())->source() . ':' . (new B())->source();
