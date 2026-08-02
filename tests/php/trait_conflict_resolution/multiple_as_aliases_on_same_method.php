<?php
// vybe-test: php/trait_conflict_resolution/multiple_as_aliases_on_same_method
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

trait Source { public function value(): int { return 42; } }
class Consumer {
    use Source {
        value as getValue;
        value as fetchValue;
    }
}
$c = new Consumer();
echo $c->getValue() . ',' . $c->fetchValue();

__vybe_check(ob_get_clean(), "42,42");
