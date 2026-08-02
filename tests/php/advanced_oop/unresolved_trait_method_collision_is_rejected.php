<?php
// vybe-test: php/advanced_oop/unresolved_trait_method_collision_is_rejected
// origin: languages/php/tests/php/test_advanced_oop.rs
// vybe-test-mode: compile

trait A { public function hello() { return "a"; } }
trait B { public function hello() { return "b"; } }
class C { use A, B; }
$c = new C();
echo $c->hello();
