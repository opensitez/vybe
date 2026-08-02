<?php
// vybe-test: php/heredoc_nowdoc/nowdoc_dollar_sign_not_interpolated
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

$price = 9.99;
$s = <<<'EOT'
Price: $price USD
EOT;
echo trim($s);

__vybe_check(ob_get_clean(), "Price: \$price USD");
