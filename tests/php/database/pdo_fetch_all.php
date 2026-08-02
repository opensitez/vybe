<?php
// vybe-test: php/database/pdo_fetch_all
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:test.db');
$stmt = $pdo->prepare('SELECT * FROM users');
$stmt->execute([]);
$rows = $stmt->fetchAll();
