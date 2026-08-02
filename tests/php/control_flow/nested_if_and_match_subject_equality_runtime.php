<?php
// vybe-test: php/control_flow/nested_if_and_match_subject_equality_runtime
// origin: languages/php/tests/php/test_control_flow.rs

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

$level = 3;
$state = match (true) {
    $level > 2 && $level < 10 => 'inner',
    default => 'outer',
};
if ($state === 'inner') {
    echo 'in';
} else {
    echo 'out';
}

__vybe_check(ob_get_clean(), "in");
