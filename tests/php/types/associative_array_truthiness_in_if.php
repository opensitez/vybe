<?php
// vybe-test: php/types/associative_array_truthiness_in_if
// origin: languages/php/tests/php/test_types.rs

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

$section = ['properties' => [], 'notes' => []]; if ($section) { echo 'yes'; } else { echo 'no'; } $section['properties']['NN'] = 'Work'; if ($section) { echo 'yes'; } else { echo 'no'; }

__vybe_check(ob_get_clean(), "yesyes");
