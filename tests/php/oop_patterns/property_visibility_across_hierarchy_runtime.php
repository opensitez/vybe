<?php
// vybe-test: php/oop_patterns/property_visibility_across_hierarchy_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Base {
    public string $public = 'pub';
    protected string $protected = 'prot';
    private string $private = 'priv';
    public function visible(): string {
        return $this->public . '|' . $this->protected;
    }
}
class Child extends Base {
    public function secret(): string {
        return $this->protected;
    }
}
$obj = new Child();
echo $obj->visible();
echo '|' . $obj->secret();
echo '|' . $obj->public;

__vybe_check(ob_get_clean(), "pub|prot|pub|prot");
