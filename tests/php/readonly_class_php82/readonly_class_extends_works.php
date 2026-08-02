<?php
// vybe-test: php/readonly_class_php82/readonly_class_extends_works
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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

readonly class Base {
    public function __construct(public string $name) {}
}
readonly class Child extends Base {
    public function __construct(string $name, public int $age) {
        parent::__construct($name);
    }
}
$c = new Child("Alice", 30);
echo $c->name . ',' . $c->age;

__vybe_check(ob_get_clean(), "Alice,30");
