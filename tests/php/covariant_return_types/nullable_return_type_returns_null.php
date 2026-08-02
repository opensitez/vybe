<?php
// vybe-test: php/covariant_return_types/nullable_return_type_returns_null
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

function maybeValue(bool $flag): ?string {
    return $flag ? "yes" : null;
}
echo maybeValue(true) . ',' . var_export(maybeValue(false), true);

__vybe_check(ob_get_clean(), "yes,NULL");
