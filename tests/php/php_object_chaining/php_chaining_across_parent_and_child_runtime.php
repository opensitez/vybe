<?php
// vybe-test: php/php_object_chaining/php_chaining_across_parent_and_child_runtime
// origin: languages/php/tests/php/test_php_object_chaining.rs

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
    public function __construct(public string $name) {}
    public function prefix(string $p): static { $this->name = $p . $this->name; return $this; }
}
class Derived extends Base {
    public function suffix(string $s): static { $this->name .= $s; return $this; }
}
$v = (new Derived('node'))->prefix('pre_')->suffix('_end');
echo $v->name;

__vybe_check(ob_get_clean(), "pre_node_end");
