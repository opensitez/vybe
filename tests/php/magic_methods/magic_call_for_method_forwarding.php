<?php
// vybe-test: php/magic_methods/magic_call_for_method_forwarding
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

class Decorator {
    public function __construct(private object $inner) {}
    public function __call($name, $args) {
        echo "before:$name ";
        $result = $this->inner->$name(...$args);
        echo "after:$name";
        return $result;
    }
}
class Service {
    public function greet(string $name): string {
        return "Hello $name";
    }
}
$d = new Decorator(new Service());
$d->greet("World");

__vybe_check(ob_get_clean(), "before:greet after:greet");
