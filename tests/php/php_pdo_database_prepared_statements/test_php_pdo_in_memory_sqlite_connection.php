<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_in_memory_sqlite_connection
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs

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

$pdo = new PDO("sqlite::memory:");
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)");
$pdo->exec("INSERT INTO users (name) VALUES ('Alice'), ('Bob')");

$stmt = $pdo->query("SELECT name FROM users ORDER BY id");
$names = $stmt->fetchAll(PDO::FETCH_COLUMN);
echo implode(", ", $names);

__vybe_check(ob_get_clean(), "Alice, Bob");
