<?php
// vybe-test: php/traits/trait_insteadof_prevents_class_parent_collision
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

trait A { public function label(): string { return 'a'; } }
trait B { public function label(): string { return 'b'; } }
class App {
    use A, B { A::label insteadof B; }
}
echo (new App())->label();

__vybe_check(ob_get_clean(), "a");
