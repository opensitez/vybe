<?php
// vybe-test: php/first_class_callables/chaining_callables_in_pipeline
// origin: languages/php/tests/php/test_first_class_callables.rs

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

function pipeline($value, callable ...$fns) {
    foreach ($fns as $fn) $value = $fn($value);
    return $value;
}
$result = pipeline(
    '  PHP is Great  ',
    trim(...),
    strtolower(...),
    fn($s) => str_replace(' ', '_', $s)
);
echo $result . "\n";

__vybe_check(ob_get_clean(), "php_is_great");
