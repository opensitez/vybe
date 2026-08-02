<?php
// vybe-test: php/database/pdo_fetch_column_offset
// origin: languages/php/tests/php/test_database.rs

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
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE items (a INTEGER, b INTEGER, c INTEGER)');
$pdo->exec('INSERT INTO items VALUES (1, 2, 3)');
$stmt = $pdo->query('SELECT a, b, c FROM items');
echo $stmt->fetchColumn(0);
echo '|';
echo $stmt->fetchColumn(2);

__vybe_check(ob_get_clean(), "1|3");
