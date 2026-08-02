<?php
// vybe-test: php/interfaces_deep/interface_chain_inheritance
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface A { public function a(): string; }
interface B extends A { public function b(): string; }
interface C extends B { public function c(): string; }
class Impl implements C {
    public function a(): string { return 'a'; }
    public function b(): string { return 'b'; }
    public function c(): string { return 'c'; }
}
$obj = new Impl();
echo $obj->a() . $obj->b() . $obj->c();
