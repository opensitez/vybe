<?php
// vybe-test: php/heredoc_nowdoc/heredoc_in_array_map_context
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

$rows = [1,2];
$labels = array_map(fn($n) => <<<EOT
item-$n
EOT, $rows);
echo implode(',', $labels);

__vybe_check(ob_get_clean(), "item-1,item-2");
