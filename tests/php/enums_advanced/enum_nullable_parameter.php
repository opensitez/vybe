<?php
// vybe-test: php/enums_advanced/enum_nullable_parameter
// origin: languages/php/tests/php/test_enums_advanced.rs

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

enum Mode { case Fast; case Safe; }
function run(?Mode $m): string { return $m === null ? 'default' : $m->name; }
echo run(null) . ',' . run(Mode::Fast);

__vybe_check(ob_get_clean(), "default,Fast");
