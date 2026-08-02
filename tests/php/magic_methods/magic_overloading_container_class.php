<?php
// vybe-test: php/magic_methods/magic_overloading_container_class
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

class PropBag {
    private array $data = [];
    public function __set($k, $v) { $this->data[$k] = $v; }
    public function __get($k) { return $this->data[$k] ?? null; }
    public function __isset($k) { return array_key_exists($k, $this->data); }
    public function __unset($k) { unset($this->data[$k]); }
    public function keys(): array { return array_keys($this->data); }
}
$bag = new PropBag();
$bag->name = "test";
$bag->value = 42;
$bag->extra = "remove me";
unset($bag->extra);
echo implode(",", $bag->keys());
echo $bag->value;
echo isset($bag->extra) ? "yes" : "no";

__vybe_check(ob_get_clean(), "name,value42no");
