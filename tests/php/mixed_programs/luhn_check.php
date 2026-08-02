<?php
// vybe-test: php/mixed_programs/luhn_check
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

function luhn(string $num): bool {
    $digits = array_reverse(str_split($num));
    $sum = 0;
    foreach ($digits as $i => $d) {
        $d = (int)$d;
        if ($i % 2 === 1) { $d *= 2; if ($d > 9) $d -= 9; }
        $sum += $d;
    }
    return $sum % 10 === 0;
}
echo luhn('4532015112830366') ? 'valid' : 'invalid';
echo ',';
echo luhn('1234567890123456') ? 'valid' : 'invalid';

__vybe_check(ob_get_clean(), "valid,invalid");
