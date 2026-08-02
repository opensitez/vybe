<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_goto_from_nested_while_in_function
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs

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

function run_with_goto(int $n): int {
    $i = 0;
    $sum = 0;
    while (true) {
        $i++;
        if ($i >= $n) {
            goto done;
        }
        $sum += $i;
    }
done:
    return $sum;
}
echo run_with_goto(4);

__vybe_check(ob_get_clean(), "6");
