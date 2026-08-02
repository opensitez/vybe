<?php
// vybe-test: php/database_prepared/prepared_execute_constraint_violation_sets_error_info
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
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_SILENT);
$pdo->exec('CREATE TABLE u (id INTEGER PRIMARY KEY)');
$pdo->exec('INSERT INTO u (id) VALUES (1)');
$stmt = $pdo->prepare('INSERT INTO u (id) VALUES (?)');
$stmt->execute([1]);
$info = $stmt->errorInfo();
echo ($info[0] ?? '') !== '00000' ? 'err' : 'ok';

__vybe_check(ob_get_clean(), "err");
