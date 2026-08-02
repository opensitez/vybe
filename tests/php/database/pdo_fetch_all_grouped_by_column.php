<?php
// vybe-test: php/database/pdo_fetch_all_grouped_by_column
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
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE items (cat TEXT, name TEXT)');
$pdo->exec("INSERT INTO items VALUES ('a', '1'), ('a', '2'), ('b', '3')");
$rows = $pdo->query('SELECT cat, name FROM items')->fetchAll(PDO::FETCH_GROUP | PDO::FETCH_ASSOC);
echo count($rows['a']);

__vybe_check(ob_get_clean(), "2");
