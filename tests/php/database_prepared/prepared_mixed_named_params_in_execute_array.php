<?php
// vybe-test: php/database_prepared/prepared_mixed_named_params_in_execute_array
// origin: languages/php/tests/php/test_database_prepared.rs

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

$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE pair (a TEXT, b TEXT)');
$stmt = $pdo->prepare('INSERT INTO pair (a, b) VALUES (:first, :second)');
$stmt->execute([':first' => 'X', ':second' => 'Y']);
$row = $pdo->query('SELECT a, b FROM pair')->fetch(PDO::FETCH_NUM);
echo $row[0] . $row[1];

__vybe_check(ob_get_clean(), "XY");
