<?php
// vybe-test: php/interfaces_deep/interface_covariant_return_runtime
// origin: languages/php/tests/php/test_interfaces_deep.rs

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
    public function __construct(public string $n) {}
}
class Child extends Base {}

interface Maker {
    public function make(string $value): Base;
}
class ChildMaker implements Maker {
    public function make(string $value): Child {
        return new Child($value);
    }
}
$maker = new ChildMaker();
$obj = $maker->make('x');
echo ($obj instanceof Child ? 'child' : 'other') . '|' . $obj->n;

__vybe_check(ob_get_clean(), "child|x");
