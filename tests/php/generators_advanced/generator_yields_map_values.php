<?php
// vybe-test: php/generators_advanced/generator_yields_map_values
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function csvRows(string $csv) {
    $lines = explode("\n", trim($csv));
    $headers = str_getcsv(array_shift($lines));
    foreach ($lines as $line) {
        $values = str_getcsv($line);
        yield array_combine($headers, $values);
    }
}
$csv = "name,age\nAlice,30\nBob,25";
foreach (csvRows($csv) as $row) {
    echo "{$row['name']} is {$row['age']}";
}

__vybe_check(ob_get_clean(), "Alice is 30Bob is 25");
