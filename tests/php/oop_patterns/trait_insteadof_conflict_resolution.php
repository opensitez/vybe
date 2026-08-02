<?php
// vybe-test: php/oop_patterns/trait_insteadof_conflict_resolution
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

trait A {
    public function hello(): string { return 'A::hello'; }
}
trait B {
    public function hello(): string { return 'B::hello'; }
}
class MyClass {
    use A, B {
        A::hello insteadof B;
        B::hello as helloFromB;
    }
}
$obj = new MyClass();
echo $obj->hello();
echo $obj->helloFromB();
