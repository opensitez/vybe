<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_getopt_short_and_long_options
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs

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

// Simulate CLI args: -f value --required=10
$_SERVER['argv'] = ['script.php', '-f', 'bar', '--required=10'];
$_SERVER['argc'] = 4;

$options = getopt("f:", ["required:"]);
echo "f=" . $options["f"] . " required=" . $options["required"];

__vybe_check(ob_get_clean(), "f= required=");
