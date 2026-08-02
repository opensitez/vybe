<?php
// vybe-test: php/database/prepared_statements
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:app.db');
$stmt = $pdo->prepare('INSERT INTO users (name, email) VALUES (?, ?)');
$stmt->execute(['Alice', 'alice@example.com']);
$stmt->execute(['Bob', 'bob@example.com']);
