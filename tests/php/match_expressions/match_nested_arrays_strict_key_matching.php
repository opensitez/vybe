<?php
// vybe-test: php/match_expressions/match_nested_arrays_strict_key_matching
// origin: languages/php/tests/php/test_match_expressions.rs

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

$payload = ['kind' => 'event', 'id' => 0];
echo match ($payload) {
    ['kind' => 'event'] => 'evt',
    ['kind' => 'log'] => 'log',
    default => 'other',
};
echo '|';
echo match ($payload['id'] ?? null) {
    0 => 'zero',
    null => 'null-id',
    default => 'other-id',
};

__vybe_check(ob_get_clean(), "other|zero");
