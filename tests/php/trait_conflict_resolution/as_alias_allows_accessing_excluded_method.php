<?php
// vybe-test: php/trait_conflict_resolution/as_alias_allows_accessing_excluded_method
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

trait A { public function hello(): string { return "A"; } }
trait B { public function hello(): string { return "B"; } }
class C {
    use A, B {
        A::hello insteadof B;
        B::hello as helloFromB;
    }
}
$c = new C();
echo $c->hello() . ',' . $c->helloFromB();

__vybe_check(ob_get_clean(), "A,B");
