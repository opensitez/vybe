<?php
// vybe-test: php/database/pdo_bind_param_by_reference_updates_on_reuse
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
$pdo->exec('CREATE TABLE audit (n INTEGER)');
$stmt = $pdo->prepare('INSERT INTO audit (n) VALUES (:n)');
$n = 4;
$stmt->bindParam(':n', $n, PDO::PARAM_INT);
$stmt->execute();
$n = 5;
$stmt->execute();
echo $pdo->query('SELECT COUNT(*) FROM audit')->fetchColumn();
echo '|';
echo $pdo->query('SELECT SUM(n) FROM audit')->fetchColumn();

__vybe_check(ob_get_clean(), "2|9");
