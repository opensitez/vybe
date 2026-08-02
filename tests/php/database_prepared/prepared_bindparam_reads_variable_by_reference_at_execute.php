<?php
// vybe-test: php/database_prepared/prepared_bindparam_reads_variable_by_reference_at_execute
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
$pdo->exec('CREATE TABLE r (id INTEGER, label TEXT)');
$pdo->exec("INSERT INTO r VALUES (1, 'one'), (2, 'two')");
$stmt = $pdo->prepare('SELECT label FROM r WHERE id = ?');
$id = 1;
$stmt->bindParam(1, $id, PDO::PARAM_INT);
$id = 2;
$stmt->execute();
echo $stmt->fetchColumn();

__vybe_check(ob_get_clean(), "two");
