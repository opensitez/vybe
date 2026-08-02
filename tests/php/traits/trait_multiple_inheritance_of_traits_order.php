<?php
// vybe-test: php/traits/trait_multiple_inheritance_of_traits_order
// origin: languages/php/tests/php/test_traits.rs

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

trait A { public function stamp(): string { return 'a'; } }
trait B { public function stamp(): string { return 'b'; } }
trait C { use A, B { A::stamp insteadof B; B::stamp as fromB; } }
class Recorder { use C; }
$r = new Recorder();
echo $r->stamp() . '|' . $r->fromB();

__vybe_check(ob_get_clean(), "a|b");
