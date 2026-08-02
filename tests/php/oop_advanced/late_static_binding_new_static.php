<?php
// vybe-test: php/oop_advanced/late_static_binding_new_static
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Animal {
    public string $name;
    public function __construct(string $name) {
        $this->name = $name;
    }
    public static function create(string $name): static {
        return new static($name);
    }
    public function type(): string { return "animal"; }
}
class Dog extends Animal {
    public function type(): string { return "dog"; }
}
$a = Animal::create("Rex");
$d = Dog::create("Buddy");
echo $a->type(), "\n";
echo $d->type(), "\n";
echo $d->name, "\n";

__vybe_check(ob_get_clean(), "animal\ndog\nBuddy");
