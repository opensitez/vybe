<?php
// vybe-test: php/database_prepared/prepared_bindvalue_named_inserts_string
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
$pdo->exec('CREATE TABLE u (name TEXT)');
$stmt = $pdo->prepare('INSERT INTO u (name) VALUES (:name)');
$stmt->bindValue(':name', 'ada', PDO::PARAM_STR);
$stmt->execute();
echo $pdo->query('SELECT name FROM u')->fetchColumn();

__vybe_check(ob_get_clean(), "ada");
