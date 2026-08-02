<?php
// vybe-test: php/programs/run_length_encoding_round_trip
// origin: languages/php/tests/php/test_programs.rs

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

function rleEncode(string $s): string {
    $out = '';
    $i = 0;
    while ($i < strlen($s)) {
        $c = $s[$i]; $cnt = 1;
        while ($i + $cnt < strlen($s) && $s[$i + $cnt] === $c) $cnt++;
        $out .= $cnt . $c;
        $i += $cnt;
    }
    return $out;
}
function rleDecode(string $s): string {
    $out = '';
    preg_match_all('/(\d+)([a-zA-Z])/', $s, $matches, PREG_SET_ORDER);
    foreach ($matches as $m) $out .= str_repeat($m[2], (int)$m[1]);
    return $out;
}
$encoded = rleEncode('AAABBBCCDDDDEE');
echo $encoded . "\n";
echo rleDecode($encoded) . "\n";

__vybe_check(ob_get_clean(), "3A3B2C4D2E\nAAABBBCCDDDDEE");
