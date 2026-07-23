use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Magic Methods — __get, __set, __isset, __unset, __call, __callStatic, __toString, __invoke, __clone, __serialize, __unserialize
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_magic_get_set_property_interception() {
    let out = run_prints(
        r#"<?php
class DynamicContainer {
    private array $storage = [];
    public function __set(string $name, mixed $value): void {
        $this->storage[$name] = $value;
    }
    public function __get(string $name): mixed {
        return $this->storage[$name] ?? null;
    }
}

$c = new DynamicContainer();
$c->foo = "bar";
echo $c->foo;
"#,
    );
    assert_eq!(out, vec!["bar"]);
}

#[test]
fn test_php_magic_call_and_callstatic_method_interception() {
    let out = run_prints(
        r#"<?php
class MagicProxy {
    public function __call(string $name, array $args): string {
        return "DYNAMIC_$name(" . implode(",", $args) . ")";
    }
    public static function __callStatic(string $name, array $args): string {
        return "STATIC_$name(" . implode(",", $args) . ")";
    }
}

$p = new MagicProxy();
echo $p->findUser(42) . " | " . MagicProxy::whereStatus("active");
"#,
    );
    assert_eq!(
        out,
        vec!["DYNAMIC_findUser(42) | STATIC_whereStatus(active)"]
    );
}

#[test]
fn test_php_magic_to_string_cast() {
    let out = run_prints(
        r#"<?php
class Money {
    public function __construct(public float $amount, public string $currency) {}
    public function __toString(): string {
        return "{$this->currency} " . number_format($this->amount, 2);
    }
}

$m = new Money(99.9, "USD");
echo "Total: $m";
"#,
    );
    assert_eq!(out, vec!["Total: USD 99.90"]);
}

#[test]
fn test_php_magic_invoke_object_as_callable() {
    let out = run_prints(
        r#"<?php
class Multiplier {
    public function __construct(private int $factor) {}
    public function __invoke(int $num): int {
        return $num * $this->factor;
    }
}

$double = new Multiplier(2);
echo $double(10);
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_php_magic_clone_deep_copy() {
    let out = run_prints(
        r#"<?php
class Address {
    public string $city = "NY";
}

class User {
    public Address $address;
    public function __construct() {
        $this->address = new Address();
    }
    public function __clone() {
        $this->address = clone $this->address;
    }
}

$u1 = new User();
$u2 = clone $u1;
$u2->address->city = "LA";
echo $u1->address->city . " vs " . $u2->address->city;
"#,
    );
    assert_eq!(out, vec!["NY vs LA"]);
}

#[test]
fn test_php_magic_serialize_unserialize_php74() {
    compile_ok(
        r#"<?php
class SessionData {
    public string $user = "Alice";
    public string $token = "secret";

    public function __serialize(): array {
        return ["u" => $this->user];
    }
    public function __unserialize(array $data): void {
        $this->user = $data["u"];
        $this->token = "guest";
    }
}

$s = new SessionData();
$str = serialize($s);
$restored = unserialize($str);
echo $restored->user . " " . $restored->token;
"#,
    );
}

#[test]
fn test_php_magic_isset_and_unset() {
    compile_ok(
        r#"<?php
class DataBag {
    private array $data = ["a" => 1];
    public function __isset(string $name): bool {
        return isset($this->data[$name]);
    }
    public function __unset(string $name): void {
        unset($this->data[$name]);
    }
}

$b = new DataBag();
echo isset($b->a) ? "YES" : "NO";
unset($b->a);
echo isset($b->a) ? "YES" : "NO";
"#,
    );
}

#[test]
fn test_php_magic_debug_info_custom_output() {
    compile_ok(
        r#"<?php
class SensitiveModel {
    private string $password = "secret123";
    public function __debugInfo(): array {
        return ["password" => "******"];
    }
}

$m = new SensitiveModel();
var_dump($m);
"#,
    );
}

#[test]
fn test_php_magic_sleep_and_wakeup_legacy() {
    compile_ok(
        r#"<?php
class Connection {
    public string $dsn = "sqlite::memory:";
    public function __sleep(): array {
        return ["dsn"];
    }
    public function __wakeup(): void {
        echo "Reconnected";
    }
}

$c = new Connection();
$s = serialize($c);
$restored = unserialize($s);
"#,
    );
}

#[test]
fn test_php_magic_set_state_export_import() {
    compile_ok(
        r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
    public static function __set_state(array $array): Point {
        return new Point($array["x"], $array["y"]);
    }
}

$p = new Point(5, 10);
eval('$p2 = ' . var_export($p, true) . ';');
echo $p2->x;
"#,
    );
}
