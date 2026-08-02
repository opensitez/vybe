<?php
// vybe-test: php/php84/property_hook_inherited
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Base {
    public int $value {
        get { return $this->value; }
        set(int $v) { $this->value = max(0, $v); }
    }
}
class Derived extends Base {
    public int $value {
        set(int $v) { $this->value = max(0, min(100, $v)); }
    }
}
$d = new Derived();
$d->value = 150;
echo $d->value;
