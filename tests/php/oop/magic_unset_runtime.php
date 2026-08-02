<?php
// vybe-test: php/oop/magic_unset_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Bag {
    private array $v = [];
    public function __set(string $name, mixed $value): void { $this->v[$name] = $value; }
    public function __get(string $name): mixed { return $this->v[$name] ?? null; }
    public function __unset(string $name): void { unset($this->v[$name]); }
    public function has(string $name): bool { return isset($this->v[$name]); }
}
$b = new Bag();
$b->x = 1;
unset($b->x);
echo $b->has('x') ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "no");
