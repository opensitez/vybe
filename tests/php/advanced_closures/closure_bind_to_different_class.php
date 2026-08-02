<?php
// vybe-test: php/advanced_closures/closure_bind_to_different_class
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class A { private string $tag = 'A'; }
class B { private string $tag = 'B'; }
$readTag = Closure::bind(function(): string { return $this->tag; }, new B(), B::class);
echo $readTag();
