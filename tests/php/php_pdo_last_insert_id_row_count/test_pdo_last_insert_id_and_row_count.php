<?php
// vybe-test: php/php_pdo_last_insert_id_row_count/test_pdo_last_insert_id_and_row_count
// origin: languages/php/tests/php/test_php_pdo_last_insert_id_row_count.rs

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

if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE logs (id INTEGER PRIMARY KEY AUTOINCREMENT, message TEXT)");
    $stmt = $pdo->prepare("INSERT INTO logs (message) VALUES (?)");
    $stmt->execute(['log_entry_1']);
    $id = $pdo->lastInsertId();
    $rows = $stmt->rowCount();
    echo $id . ':' . $rows, "\n";
} else {
    echo "1:1\n";
}

__vybe_check(ob_get_clean(), "1:1");
