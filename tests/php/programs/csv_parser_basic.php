<?php
// vybe-test: php/programs/csv_parser_basic
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

function parseCsv(string $input): array {
    return array_map(fn($line) => explode(',', $line), explode("\n", trim($input)));
}
$csv = "name,age,city\nAlice,30,NYC\nBob,25,LA";
$rows = parseCsv($csv);
echo count($rows) . "\n";
echo $rows[0][0] . "\n";
echo $rows[1][1] . "\n";
echo $rows[2][2] . "\n";

__vybe_check(ob_get_clean(), "3\nname\n30\nLA");
