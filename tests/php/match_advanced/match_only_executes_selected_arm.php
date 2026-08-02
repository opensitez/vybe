<?php
// vybe-test: php/match_advanced/match_only_executes_selected_arm
// origin: languages/php/tests/php/test_match_advanced.rs

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

function side_effect(string &$log, string $label): string {
    $log .= $label;
    return $label;
}
$log = '';
$value = 2;
$out = match ($value) {
    1 => side_effect($log, 'A'),
    2 => 'selected',
    default => side_effect($log, 'Z'),
};
echo $out;
echo '|';
echo $log;

__vybe_check(ob_get_clean(), "selected|");
