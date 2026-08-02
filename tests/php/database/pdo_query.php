<?php
// vybe-test: php/database/pdo_query
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:test.db');
$rows = $pdo->query('SELECT * FROM users');
