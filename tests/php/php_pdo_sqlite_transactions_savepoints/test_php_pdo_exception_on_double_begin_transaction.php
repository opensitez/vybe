<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_exception_on_double_begin_transaction
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->beginTransaction();

try {
    $pdo->beginTransaction(); // Double begin transaction throws exception
} catch (PDOException $e) {
    echo "Double beginTransaction caught: " . $e->getMessage();
}
$pdo->rollBack();
