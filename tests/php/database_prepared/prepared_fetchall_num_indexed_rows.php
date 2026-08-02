<?php
// vybe-test: php/database_prepared/prepared_fetchall_num_indexed_rows
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
$pdo->exec('CREATE TABLE p (a INTEGER, b INTEGER)');
$pdo->exec('INSERT INTO p VALUES (1, 10), (2, 20)');
$stmt = $pdo->prepare('SELECT a, b FROM p ORDER BY a');
$stmt->execute();
$rows = $stmt->fetchAll(PDO::FETCH_NUM);
echo $rows[1][1];

__vybe_check(ob_get_clean(), "20");
