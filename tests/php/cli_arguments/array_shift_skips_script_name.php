<?php
// vybe-test: php/cli_arguments/array_shift_skips_script_name
// origin: languages/php/tests/php/test_cli_arguments.rs

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

$argv = ['run.php', 'build', 'release'];
array_shift($argv);
echo implode(',', $argv);

__vybe_check(ob_get_clean(), "build,release");
