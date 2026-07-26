use super::helpers::run_prints;

#[test]
fn test_property_hooks_get_set_syntax() {
    assert_eq!(
        run_prints(
            r#"<?php
class UserHookDemo {
    private string $_first = 'John';
    private string $_last = 'Doe';

    public string $fullName {
        get => $this->_first . ' ' . $this->_last;
    }
}
$u = new UserHookDemo();
echo $u->fullName, "\n";
"#
        ),
        vec!["John Doe"]
    );
}

#[test]
fn test_property_hooks_set_transformation() {
    assert_eq!(
        run_prints(
            r#"<?php
class SanitizedUser {
    private string $_email = '';

    public string $email {
        get => $this->_email;
        set => $this->_email = strtolower(trim($value));
    }
}
$u = new SanitizedUser();
$u->email = '   USER@EXAMPLE.COM   ';
echo $u->email, "\n";
"#
        ),
        vec!["user@example.com"]
    );
}

#[test]
fn test_property_hooks_get_set_runtime_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
class Product {
    private float $_price = 0.0;

    public float $price {
        get => $this->_price;
        set => $this->_price = $value;
    }
}
$p = new Product();
$p->price = 19.95;
echo number_format($p->price, 2);
"#,
        ),
        vec!["19.95"]
    );
}

#[test]
fn test_property_hooks_block_setter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Logger {
    private string $_prefix = '';

    public string $tag {
        get => $this->_prefix;
        set {
            $this->_prefix = strtoupper($value);
        }
    }
}
$x = new Logger();
$x->tag = 'dev';
echo $x->tag;
"#,
        ),
        vec!["DEV"]
    );
}

#[test]
fn test_property_hooks_setter_with_default_when_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
class ApiToken {
    private string $token = '';
    public string $publicToken {
        get => $this->token === '' ? 'unset' : $this->token;
        set {
            $this->token = trim($value) === '' ? 'default' : trim($value);
        }
    }
}
$t = new ApiToken();
$t->publicToken = ' ';
echo $t->publicToken;
$t->publicToken = 'abc123';
echo '|' . $t->publicToken;
"#,
        ),
        vec!["default|abc123"]
    );
}

#[test]
fn test_property_hooks_setter_type_enforced_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class CounterState {
    private int $value = 0;
    public int $count {
        get => $this->value;
        set => $this->value = max(0, $value);
    }
}
$c = new CounterState();
$c->count = -7;
echo $c->count . '|';
$c->count = 7;
echo $c->count;
"#,
        ),
        vec!["0|7"]
    );
}
