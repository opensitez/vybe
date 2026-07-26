use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

#[test]
fn php_method_chaining_with_ternary_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Calculator {
    public function __construct(public int $value) {}
    public function inc(int $n): self { $this->value += $n; return $this; }
    public function dec(int $n): self { $this->value -= $n; return $this; }
    public function done(): int { return $this->value; }
}
$c = (new Calculator(3))->inc(7)->dec(4);
echo $c->done();
"#,
        ),
        vec!["6"]
    );
}

#[test]
fn php_method_chaining_runtime() {
    for n in 1..=20_i64 {
        let expected = n * n;

        assert_int(
            &format!(
                "class ChainValue {{\n    public int $value;\n    public function __construct(int $v) {{ $this->value = $v; }}\n    public function add(int $v): self {{ $this->value += $v; return $this; }}\n    public function mul(int $v): self {{ $this->value *= $v; return $this; }}\n    public function sub(int $v): self {{ $this->value -= $v; return $this; }}\n    public function val(): int {{ return $this->value; }}\n}}\n\necho (new ChainValue({n}))->add({n})->mul({n})->sub({n})->val();",
                n = n,
            ),
            expected,
        );
    }
}

#[test]
fn php_fluid_setters_return_this_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Profile {
    public string $name = '';
    public string $role = '';
    public function set_name(string $name): self { $this->name = $name; return $this; }
    public function set_role(string $role): self { $this->role = $role; return $this; }
    public function label(): string { return $this->name . ':' . $this->role; }
}
$p = (new Profile())->set_name('alice')->set_role('admin');
echo $p->label();
"#,
        ),
        vec!["alice:admin"]
    );
}

#[test]
fn php_chaining_across_parent_and_child_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public function __construct(public string $name) {}
    public function prefix(string $p): static { $this->name = $p . $this->name; return $this; }
}
class Derived extends Base {
    public function suffix(string $s): static { $this->name .= $s; return $this; }
}
$v = (new Derived('node'))->prefix('pre_')->suffix('_end');
echo $v->name;
"#,
        ),
        vec!["pre_node_end"]
    );
}

#[test]
fn php_magic_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public function __construct(private int $value) {}
    public function inc(): self { $this->value += 1; return $this; }
    public function __get(string $name): mixed { return $this->value; }
}
$c = (new Counter(0))->inc()->inc();
echo $c->value;
"#,
        ),
        vec!["2"]
    );
}

#[test]
fn php_chain_constructor_to_static_factory_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Logger {
    public function __construct(public string $value) {}
    public static function from(string $value): static { return new static($value); }
    public function append(string $suffix): static { $this->value .= $suffix; return $this; }
}
$v = Logger::from('a')->append('b')->append('c');
echo $v->value;
"#,
        ),
        vec!["abc"]
    );
}
