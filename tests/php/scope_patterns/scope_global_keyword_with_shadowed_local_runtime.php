<?php
// vybe-test: php/scope_patterns/scope_global_keyword_with_shadowed_local_runtime
// origin: languages/php/tests/php/test_scope_patterns.rs

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

$counter = 10;
function increment_global(): int {
    global $counter;
    $counter += 5;
    return $counter;
}
function read_local_counter(): int {
    $counter = 2;
    return $counter;
}
echo read_local_counter();
echo '|';
echo increment_global();
echo '|';
echo increment_global();

__vybe_check(ob_get_clean(), "2|15|20");
