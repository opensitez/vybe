<?php
// vybe-test: php/database/query_with_loop
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:app.db');
$rows = $pdo->query('SELECT * FROM users');
foreach ($rows as $row) {
    echo $row['name'] . ': ' . $row['email'];
}
