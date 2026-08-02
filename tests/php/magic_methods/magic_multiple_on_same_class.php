<?php
// vybe-test: php/magic_methods/magic_multiple_on_same_class
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

class SmartObj {
    private array $props = [];
    public function __get($k) { return $this->props[$k] ?? null; }
    public function __set($k, $v) { $this->props[$k] = $v; }
    public function __isset($k) { return isset($this->props[$k]); }
    public function __toString(): string { return json_encode($this->props); }
    public function __invoke() { return count($this->props); }
}
$s = new SmartObj();
$s->x = 1;
$s->y = 2;
echo $s->x;
echo isset($s->y) ? "yes" : "no";
echo $s();
echo $s;

__vybe_check(ob_get_clean(), "1yes2{\"x\":1,\"y\":2}");
