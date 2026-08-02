<?php
// vybe-test: php/method_chaining/chain_static_factory_then_instance_methods
// origin: languages/php/tests/php/test_method_chaining.rs

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

class Config {
    private array $data = [];
    public static function make(): static { return new static(); }
    public function set(string $k, mixed $v): static { $this->data[$k] = $v; return $this; }
    public function get(string $k): mixed { return $this->data[$k] ?? null; }
}
echo Config::make()->set('port', 8080)->get('port');

__vybe_check(ob_get_clean(), "8080");
