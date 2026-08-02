<?php
// vybe-test: php/database/pdo_rollback
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("DELETE FROM users");
$pdo->rollBack();
