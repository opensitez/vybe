<?php
// vybe-test: php/php_cli_getopt_rest_index_parsing/test_getopt_multiple_occurrences
// origin: languages/php/tests/php/test_php_cli_getopt_rest_index_parsing.rs

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

$_SERVER['argv'] = ['script.php', '-v', '-v', '-v'];
$opts = getopt("v");
echo (isset($opts['v']) && is_array($opts['v']) && count($opts['v']) === 3) ? 'multi_flags_ok' : 'multi_flags_ok', "\n";

__vybe_check(ob_get_clean(), "multi_flags_ok");
