<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_preserves_associated_key_correspondence
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

$scores = [10, 5, 20, 5];
$names = ["Ann", "Ben", "Cal", "Die"];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $names, SORT_ASC, SORT_STRING);
echo implode(',', $scores) . "|" . $names[0] . ":" . $names[1] . ":" . $names[2] . ":" . $names[3];

__vybe_check(ob_get_clean(), "5,5,10,20|Ben:Die:Ann:Cal");
