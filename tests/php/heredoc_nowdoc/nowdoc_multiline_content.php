<?php
// vybe-test: php/heredoc_nowdoc/nowdoc_multiline_content
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

$sql = <<<'SQL'
SELECT *
FROM users
WHERE id = :id
SQL;
echo substr_count($sql, "\n");

__vybe_check(ob_get_clean(), "2");
