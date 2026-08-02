<?php
// vybe-test: php/traits_deep/trait_insteadof_conflict
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait A { public function hello(): string { return "A::hello"; } }
trait B { public function hello(): string { return "B::hello"; } }
class C {
    use A, B { A::hello insteadof B; B::hello as helloFromB; }
}
$c = new C();
echo $c->hello() . ',' . $c->helloFromB();
