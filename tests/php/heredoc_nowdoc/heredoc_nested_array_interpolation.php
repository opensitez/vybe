<?php
// vybe-test: php/heredoc_nowdoc/heredoc_nested_array_interpolation
// origin: languages/php/tests/php/test_heredoc_nowdoc.rs

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

$users = [['name' => 'Alice'], ['name' => 'Bob']];
$s = <<<EOT
First: {$users[0]['name']}
EOT;
echo trim($s);

__vybe_check(ob_get_clean(), "First: Alice");
