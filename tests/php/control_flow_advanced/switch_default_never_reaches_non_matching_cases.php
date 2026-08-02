<?php
// vybe-test: php/control_flow_advanced/switch_default_never_reaches_non_matching_cases
// origin: languages/php/tests/php/test_control_flow_advanced.rs

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

$state = 'z';
$out = [];
switch ($state) {
    case 'a':
        $out[] = 'a';
        break;
    default:
        $out[] = 'd';
        break;
    case 'b':
        $out[] = 'b';
        break;
}
echo implode('|', $out);

__vybe_check(ob_get_clean(), "d");
