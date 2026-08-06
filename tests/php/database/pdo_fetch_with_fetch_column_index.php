<?php
// vybe-test: php/database/pdo_fetch_with_fetch_column_index
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
$pdo->exec('CREATE TABLE t (x INTEGER, y INTEGER)');
$pdo->exec('INSERT INTO t VALUES (11, 22), (33, 44)');
$stmt = $pdo->prepare('SELECT x, y FROM t ORDER BY x');
$stmt->execute();
echo $stmt->fetchColumn(1);
echo '|';
echo $stmt->fetchColumn(1);

__vybe_check(ob_get_clean(), "22|44");
