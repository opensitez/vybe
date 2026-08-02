<?php
// vybe-test: php/magic_methods/magic_clone_original_unchanged
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

class Node {
    public function __construct(public string $value, public ?Node $next = null) {}
    public function __clone() {
        if ($this->next !== null) {
            $this->next = clone $this->next;
        }
    }
}
$a = new Node("first", new Node("second"));
$b = clone $a;
$b->value = "modified";
$b->next->value = "also modified";
echo $a->value;
echo $a->next->value;
echo $b->value;
echo $b->next->value;

__vybe_check(ob_get_clean(), "firstsecondmodifiedalso modified");
