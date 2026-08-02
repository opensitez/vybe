<?php
// vybe-test: php/oop_advanced/magic_property_hooks_runtime_check
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

class Box {
    private array $store = [];
    public function __set(string $name, mixed $value): void { $this->store[$name] = $value; }
    public function __get(string $name): mixed { return $this->store[$name] ?? null; }
    public function has(string $name): bool { return array_key_exists($name, $this->store); }
}
$b = new Box();
$b->x = 10;
echo $b->has('x') ? 'yes' : 'no';
echo $b->x;

__vybe_check(ob_get_clean(), "yes10");
