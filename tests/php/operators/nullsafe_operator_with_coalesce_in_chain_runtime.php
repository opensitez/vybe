<?php
// vybe-test: php/operators/nullsafe_operator_with_coalesce_in_chain_runtime
// origin: languages/php/tests/php/test_operators.rs

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

class Child {
    public string $value = 'ok';
    public function getChild(): ?Child { return null; }
}
class ParentObj {
    public Child $child;
    public function __construct() { $this->child = new Child(); }
}
$obj = new ParentObj();
echo $obj->child?->value . '|';
echo $obj->child->getChild()?->value ?? 'none';

__vybe_check(ob_get_clean(), "ok|none");
