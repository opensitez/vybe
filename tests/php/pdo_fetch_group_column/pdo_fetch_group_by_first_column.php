<?php
// vybe-test: php/pdo_fetch_group_column/pdo_fetch_group_by_first_column
// origin: languages/php/tests/php/test_pdo_fetch_group_column.rs

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
$pdo->exec("CREATE TABLE colors (group_id INTEGER, color TEXT)");
$pdo->exec("INSERT INTO colors VALUES (1, 'red')");
$pdo->exec("INSERT INTO colors VALUES (1, 'blue')");
$pdo->exec("INSERT INTO colors VALUES (2, 'green')");

$stmt = $pdo->query("SELECT group_id, color FROM colors");
$grouped = $stmt->fetchAll(PDO::FETCH_COLUMN | PDO::FETCH_GROUP);

echo count($grouped[1]) . "|" . $grouped[1][0] . "|" . $grouped[2][0];

__vybe_check(ob_get_clean(), "2|red|green");
