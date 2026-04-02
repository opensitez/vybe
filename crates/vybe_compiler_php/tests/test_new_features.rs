mod helpers;
use helpers::compile_ok;

// ── String interpolation ────────────────────────────────────
#[test] fn interpolation_simple() { compile_ok(r#"<?php $name = "World"; echo "Hello $name";"#); }
#[test] fn interpolation_curly() { compile_ok(r#"<?php $name = "World"; echo "Hello {$name}!";"#); }
#[test] fn interpolation_multiple() { compile_ok(r#"<?php $a = "foo"; $b = "bar"; echo "$a and $b";"#); }
#[test] fn interpolation_in_middle() { compile_ok(r#"<?php $x = 42; echo "The answer is $x ok";"#); }
#[test] fn no_interpolation_single() { compile_ok("<?php echo 'no $interpolation here';"); }
#[test] fn escaped_dollar() { compile_ok(r#"<?php echo "price is \$5";"#); }

// ── Closure use captures ────────────────────────────────────
#[test] fn closure_use() { compile_ok("<?php $x = 10; $fn = function() use ($x) { return $x; }; echo $fn();"); }
#[test] fn closure_use_multiple() { compile_ok("<?php $a = 1; $b = 2; $fn = function() use ($a, $b) { return $a + $b; }; echo $fn();"); }
#[test] fn arrow_fn_captures() { compile_ok("<?php $x = 5; $fn = fn($y) => $x + $y; echo $fn(3);"); }

// ── Null coalesce assignment ────────────────────────────────
#[test] fn null_coalesce_assign() { compile_ok("<?php $x = null; $x ??= 'default'; echo $x;"); }
#[test] fn null_coalesce_assign_non_null() { compile_ok("<?php $x = 'existing'; $x ??= 'default'; echo $x;"); }

// ── List / array destructuring ──────────────────────────────
#[test] fn list_assign() { compile_ok("<?php list($a, $b) = [1, 2]; echo $a;"); }
#[test] fn short_list_assign() { compile_ok("<?php [$a, $b, $c] = [10, 20, 30]; echo $b;"); }

// ── Heredoc / Nowdoc ────────────────────────────────────────
#[test] fn heredoc_basic() { compile_ok("<?php $x = <<<EOT\nHello World\nEOT;\necho $x;"); }
#[test] fn nowdoc_basic() { compile_ok("<?php $x = <<<'EOT'\nHello World\nEOT;\necho $x;"); }

// ── Traits ──────────────────────────────────────────────────
#[test] fn trait_basic() { compile_ok(r#"<?php
trait Greetable {
    public function greet() { return "Hello from " . $this->name; }
}
class Person {
    use Greetable;
    public $name;
    public function __construct($name) { $this->name = $name; }
}
$p = new Person("John");
echo $p->greet();
"#); }

#[test] fn trait_multiple() { compile_ok(r#"<?php
trait HasName {
    public function getName() { return $this->name; }
}
trait HasAge {
    public function getAge() { return $this->age; }
}
class User {
    use HasName;
    use HasAge;
    public $name;
    public $age;
    public function __construct($name, $age) { $this->name = $name; $this->age = $age; }
}
$u = new User("Alice", 30);
"#); }

// ── parent:: calls ──────────────────────────────────────────
#[test] fn parent_method_call() { compile_ok(r#"<?php
class Base {
    public function hello() { return "Hello from Base"; }
}
class Child extends Base {
    public function hello() { return parent::hello() . " and Child"; }
}
$c = new Child();
"#); }

// ── Static property access ──────────────────────────────────
#[test] fn static_const_access() { compile_ok("<?php class Cfg { const VER = '1.0'; } echo Cfg::VER;"); }
#[test] fn static_method_call() { compile_ok("<?php class M { public static function sq($n) { return $n * $n; } } echo M::sq(5);"); }

// ── Interfaces ──────────────────────────────────────────────
#[test] fn interface_decl() { compile_ok(r#"<?php
interface Printable {
    public function toString();
}
class Item implements Printable {
    public $name;
    public function __construct($n) { $this->name = $n; }
    public function toString() { return $this->name; }
}
$i = new Item("test");
echo $i->toString();
"#); }

// ── Type juggling ───────────────────────────────────────────
#[test] fn string_plus_number() { compile_ok("<?php $x = '5' + 3; echo $x;"); }
#[test] fn string_sub() { compile_ok("<?php $x = '10' - 3; echo $x;"); }
#[test] fn string_mul() { compile_ok("<?php $x = '4' * 5; echo $x;"); }
#[test] fn bool_arithmetic() { compile_ok("<?php $x = true + true; echo $x;"); }
