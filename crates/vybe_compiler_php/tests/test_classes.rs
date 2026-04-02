mod helpers;
use helpers::compile_ok;

#[test] fn class_empty() { compile_ok("<?php class Foo {} $f = new Foo();"); }
#[test] fn class_constructor() { compile_ok("<?php class Dog { public $name; public function __construct($name) { $this->name = $name; } } $d = new Dog('Rex');"); }
#[test] fn class_method() { compile_ok("<?php class Dog { public $name; public function __construct($n) { $this->name = $n; } public function speak() { return $this->name . ' says Woof'; } } $d = new Dog('Rex'); echo $d->speak();"); }
#[test] fn class_inheritance() { compile_ok("<?php class Animal { public $name; public function __construct($n) { $this->name = $n; } } class Cat extends Animal { public function speak() { return $this->name . ' says Meow'; } } $c = new Cat('Whiskers'); echo $c->speak();"); }
#[test] fn class_property_default() { compile_ok("<?php class Config { public $debug = false; public $version = '1.0'; } $c = new Config();"); }
#[test] fn static_method() { compile_ok("<?php class MathHelper { public static function square($n) { return $n * $n; } } echo MathHelper::square(5);"); }
#[test] fn class_constant() { compile_ok("<?php class Config { const VERSION = '1.0'; } echo Config::VERSION;"); }
#[test] fn multiple_methods() { compile_ok("<?php class Calc { public function add($a,$b) { return $a+$b; } public function sub($a,$b) { return $a-$b; } } $c = new Calc(); echo $c->add(3,2);"); }
#[test] fn chained_calls() { compile_ok("<?php class Builder { public $val = ''; public function add($s) { $this->val .= $s; return $this; } } $b = new Builder(); $b->add('a')->add('b');"); }
#[test] fn new_with_args() { compile_ok("<?php class Point { public $x; public $y; public function __construct($x, $y) { $this->x = $x; $this->y = $y; } } $p = new Point(1, 2);"); }
