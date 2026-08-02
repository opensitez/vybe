<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_sqlite_user_defined_function
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
if (method_exists($pdo, "sqliteCreateFunction")) {
    $pdo->sqliteCreateFunction("md5_hash", fn($val) => md5($val));
    $stmt = $pdo->query("SELECT md5_hash('test')");
    echo "Hash: " . $stmt->fetchColumn();
} else {
    echo "sqliteCreateFunction not supported";
}
