<?php
// vybe-test: php/mixed_programs/run_length_encoding
// origin: languages/php/tests/php/test_mixed_programs.rs

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

function rle(string $s): string {
    $result = '';
    $i = 0;
    while ($i < strlen($s)) {
        $c = $s[$i]; $count = 1;
        while ($i + $count < strlen($s) && $s[$i + $count] === $c) $count++;
        $result .= $count > 1 ? $count . $c : $c;
        $i += $count;
    }
    return $result;
}
echo rle('AAABBBCCDDDDEE');

__vybe_check(ob_get_clean(), "3A3B2C4D2E");
