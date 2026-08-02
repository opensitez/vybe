<?php
// vybe-test: php/database/crud_pattern
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:app.db');
$pdo->exec('CREATE TABLE IF NOT EXISTS posts (id INTEGER PRIMARY KEY, title TEXT, body TEXT)');

// Create
$pdo->exec("INSERT INTO posts (title, body) VALUES ('Hello', 'World')");

// Read
$posts = $pdo->query('SELECT * FROM posts');

// Update
$pdo->exec("UPDATE posts SET title = 'Updated' WHERE id = 1");

// Delete
$pdo->exec('DELETE FROM posts WHERE id = 1');
