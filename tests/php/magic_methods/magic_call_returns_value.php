<?php
// vybe-test: php/magic_methods/magic_call_returns_value
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

class Accessor {
    private array $data = ["name" => "John", "age" => 30];
    public function __call($method, $args) {
        if (str_starts_with($method, "get")) {
            $prop = strtolower(substr($method, 3));
            return $this->data[$prop] ?? null;
        }
        return null;
    }
}
$a = new Accessor();
echo $a->getName();
echo $a->getAge();

__vybe_check(ob_get_clean(), "John30");
