<?php
// vybe-test: php/exception_types/bad_method_call_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

class Foo {
    public function bar(): void {
        throw new BadMethodCallException('bar not implemented');
    }
}
try { (new Foo())->bar(); } catch (BadMethodCallException $e) { echo $e->getMessage(); }
