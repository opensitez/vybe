<?php
// vybe-test: php/clone_patterns/parent_clone_called_from_child_clone
// origin: languages/php/tests/php/test_clone_patterns.rs

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
    public array $tags = [];
    public function __clone() { $this->tags[] = 'base_cloned'; }
}
class Child extends Base {
    public function __clone() {
        parent::__clone();
        $this->tags[] = 'child_cloned';
    }
}
$c = new Child();
$d = clone $c;
echo implode(',', $d->tags);

__vybe_check(ob_get_clean(), "base_cloned,child_cloned");
