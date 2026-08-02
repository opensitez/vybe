<?php
// vybe-test: php/magic_methods/magic_method_in_abstract_class
// origin: languages/php/tests/php/test_magic_methods.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

abstract class BaseEntity {
    protected array $attributes = [];
    public function __get($name) { return $this->attributes[$name] ?? null; }
    public function __set($name, $value) { $this->attributes[$name] = $value; }
    abstract public function getType(): string;
}
class User extends BaseEntity {
    public function getType(): string { return "user"; }
}
$u = new User();
$u->name = "Alice";
echo $u->name;
echo $u->getType();

__vybe_check(ob_get_clean(), "Aliceuser");
