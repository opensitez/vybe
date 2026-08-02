<?php
// vybe-test: php/string_interpolation_complex/interpolation_heredoc_vs_nowdoc_behavior_runtime
// origin: languages/php/tests/php/test_string_interpolation_complex.rs

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

$name = "Alice";
$doc = <<<TXT
Hello $name
TXT;

$raw = <<<'TXT'
Hello $name
TXT;

echo $doc;
echo "\n";
echo $raw;

__vybe_check(ob_get_clean(), "Hello Alice|Hello \$name");
