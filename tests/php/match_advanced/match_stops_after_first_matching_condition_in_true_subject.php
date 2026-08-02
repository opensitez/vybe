<?php
// vybe-test: php/match_advanced/match_stops_after_first_matching_condition_in_true_subject
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

$value = 10;
$log = '';
function mark(array &$log, string $value): string {
    $log[] = $value;
    return $value;
}
$label = match (true) {
    $value > 5 && mark($log, 'high') => 'high',
    $value > 0 && mark($log, 'positive') => 'positive',
    default => 'zero',
};
echo $label;
echo '|';
echo implode(',', $log);

__vybe_check(ob_get_clean(), "high|high");
