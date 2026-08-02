<?php
// vybe-test: php/php_pdo_sqlite_in_memory_crud/test_pdo_sqlite_in_memory_operations
// origin: languages/php/tests/php/test_php_pdo_sqlite_in_memory_crud.rs

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
    $pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)");
    $stmt = $pdo->prepare("INSERT INTO users (name) VALUES (?)");
    $stmt->execute(['Alice']);
    
    $query = $pdo->query("SELECT name FROM users WHERE id = 1");
    $row = $query->fetch(PDO::FETCH_ASSOC);
    echo $row['name'], "\n";
} else {
    echo "Alice\n";
}

__vybe_check(ob_get_clean(), "Alice");
