<?php
// vybe-test: php/traits_deep/trait_visibility_as
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Greetable {
    public function hello(): string { return "Hello!"; }
    public function goodbye(): string { return "Goodbye!"; }
}
class Formal {
    use Greetable { goodbye as protected; }
    public function farewell(): string { return $this->goodbye(); }
}
$f = new Formal();
echo $f->hello();
echo $f->farewell();
