<?php
// vybe-test: php/magic_methods/magic_wakeup_restores_state
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

class Cache {
    public string $key = "my_key";
    private array $data = [];
    public function set(string $k, mixed $v): void { $this->data[$k] = $v; }
    public function get(string $k): mixed { return $this->data[$k] ?? null; }
    public function __sleep(): array { return ["key"]; }
    public function __wakeup(): void { $this->data = []; }
}
$c = new Cache();
$c->set("a", 42);
$raw = serialize($c);
$c2 = unserialize($raw);
echo $c2->key;
echo $c2->get("a") === null ? "cleared" : "kept";

__vybe_check(ob_get_clean(), "my_keycleared");
