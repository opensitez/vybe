<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_get_set_property_interception
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs

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

__vybe_check(ob_get_clean(), "bar");
