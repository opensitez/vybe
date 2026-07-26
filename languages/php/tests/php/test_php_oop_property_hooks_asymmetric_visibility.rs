use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: PHP 8.4 Property Hooks & Asymmetric Visibility — get, set, backing field, public private(set), protected(set)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_property_hooks_virtual_property() {
    let out = run_prints(
        r#"<?php
class Rectangle {
    public function __construct(public float $width, public float $height) {}

    public float $area {
        get => $this->width * $this->height;
    }
}

$r = new Rectangle(5.0, 4.0);
echo $r->area;
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_php84_property_hooks_custom_setter_transformation() {
    let out = run_prints(
        r#"<?php
class UserProfile {
    public string $email {
        set => strtolower(trim($value));
    }
}

$u = new UserProfile();
$u->email = "  ALICE@Example.COM  ";
echo $u->email;
"#,
    );
    assert_eq!(out, vec!["alice@example.com"]);
}

#[test]
fn test_php84_asymmetric_visibility_public_get_private_set() {
    let out = run_prints(
        r#"<?php
class BankAccount {
    public private(set) float $balance = 0.0;

    public function deposit(float $amount): void {
        $this->balance += $amount;
    }
}

$account = new BankAccount();
$account->deposit(250.0);
echo "Balance: {$account->balance}";
"#,
    );
    assert_eq!(out, vec!["Balance: 250"]);
}

#[test]
fn test_php84_property_hooks_interface_requirement() {
    assert_eq!(
        run_prints(
        r#"<?php
interface Named {
    public string $name { get; }
}

class Company implements Named {
    public function __construct(public string $name) {}
}

$c = new Company("Acme Corp");
echo $c->name;
"#,
        ),
        vec!["Acme Corp"]
    );
}

#[test]
fn test_php84_asymmetric_visibility_protected_set() {
    assert_eq!(
        run_prints(
        r#"<?php
class BaseDocument {
    public protected(set) string $title = "Untitled";
}

class ArticleDocument extends BaseDocument {
    public function setTitle(string $title): void {
        $this->title = $title;
    }
}

$art = new ArticleDocument();
$art->setTitle("PHP 8.4 Released");
echo $art->title;
"#,
        ),
        vec!["PHP 8.4 Released"]
    );
}

#[test]
fn test_php84_property_hooks_get_set_block_body() {
    assert_eq!(
        run_prints(
        r#"<?php
class Counter {
    private int $count = 0;

    public int $value {
        get {
            return $this->count;
        }
        set {
            if ($value < 0) {
                throw new InvalidArgumentException("Count cannot be negative");
            }
            $this->count = $value;
        }
    }
}

$c = new Counter();
$c->value = 10;
echo $c->value;
"#,
        ),
        vec!["10"]
    );
}

#[test]
fn test_php84_property_hooks_abstract_property_in_abstract_class() {
    compile_ok(
        r#"<?php
abstract class Widget {
    abstract public string $label { get; }
}

class ButtonWidget extends Widget {
    public string $label { get => "Click Me"; }
}

$b = new ButtonWidget();
echo $b->label;
"#,
    );
}

#[test]
fn test_php84_asymmetric_visibility_readonly_combination() {
    compile_ok(
        r#"<?php
class Token {
    public private(set) readonly string $hash;

    public function __construct(string $secret) {
        $this->hash = md5($secret);
    }
}

$t = new Token("my_secret");
echo $t->hash;
"#,
    );
}

#[test]
fn test_php84_property_hooks_by_ref_getter() {
    compile_ok(
        r#"<?php
class Matrix {
    private array $data = [1, 2, 3];

    public array &$items {
        &get => $this->data;
    }
}

$m = new Matrix();
$items = &$m->items;
$items[0] = 99;
echo $m->items[0];
"#,
    );
}

#[test]
fn test_php84_property_hooks_final_hook() {
    compile_ok(
        r#"<?php
class SecureModel {
    public string $id {
        final get => "SECURE_ID";
    }
}

$sm = new SecureModel();
echo $sm->id;
"#,
    );
}

#[test]
fn test_php84_private_set_blocked_from_outside_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Ledger {
    public private(set) int $balance = 0;
}
$l = new Ledger();
try {
    $l->balance = 10;
    echo 'wrote';
} catch (Error $e) {
    echo 'error';
}
"#,
        ),
        vec!["error"]
    );
}

#[test]
fn test_php84_get_set_hooks_transform_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class ScoreBoard {
    private int $value = 1;

    public int $valueHooked {
        get => $this->value * 3;
        set => max(0, $value);
    }
}
$board = new ScoreBoard();
$board->valueHooked = 4;
echo $board->valueHooked;
"#,
        ),
        vec!["12"]
    );
}
