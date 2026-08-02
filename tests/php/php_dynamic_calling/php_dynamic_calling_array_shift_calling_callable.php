<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_array_shift_calling_callable
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

class Shifter {
    public function inc(int $n): int { return $n + 1; }
}
$target = [new Shifter(), 'inc'];
$first = array_shift($target);
echo is_object($first) ? 'obj' : 'no';
echo is_string($target[0]) ? 'fn' : 'bad';

__vybe_check(ob_get_clean(), "objfn");
