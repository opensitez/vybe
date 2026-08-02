<?php
// vybe-test: php/host_extra/datetime_db_pattern
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:app.db');
$now = new DateTime();
$timestamp = $now->format('Y-m-d H:i:s');
$pdo->exec("INSERT INTO logs (created_at) VALUES ('" . $timestamp . "')");
