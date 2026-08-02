<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_sqlite_create_function
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->sqliteCreateFunction("php_md5", fn($str) => md5($str));

$stmt = $pdo->query("SELECT php_md5('test')");
echo $stmt->fetchColumn();
