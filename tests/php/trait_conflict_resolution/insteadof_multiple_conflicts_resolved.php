<?php
// vybe-test: php/trait_conflict_resolution/insteadof_multiple_conflicts_resolved
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs

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

trait X { public function foo(): string { return "X:foo"; } public function bar(): string { return "X:bar"; } }
trait Y { public function foo(): string { return "Y:foo"; } public function bar(): string { return "Y:bar"; } }
class Z {
    use X, Y {
        X::foo insteadof Y;
        Y::bar insteadof X;
    }
}
$z = new Z();
echo $z->foo() . ',' . $z->bar();

__vybe_check(ob_get_clean(), "X:foo,Y:bar");
