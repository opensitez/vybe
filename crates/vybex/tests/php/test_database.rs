use super::helpers::{compile_ok, run_prints};

// ── PDO style ───────────────────────────────────────────────
#[test] fn pdo_connect_sqlite() { compile_ok(r#"<?php $pdo = new PDO('sqlite:test.db');"#); }
#[test] fn pdo_connect_mysql() { compile_ok(r#"<?php $pdo = new PDO('mysql:host=localhost;dbname=mydb');"#); }
#[test] fn pdo_connect_postgres() { compile_ok(r#"<?php $pdo = new PDO('postgresql://localhost/mydb');"#); }

#[test] fn pdo_query() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$rows = $pdo->query('SELECT * FROM users');
"#); }

#[test] fn pdo_exec() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->exec('CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT)');
$pdo->exec("INSERT INTO users (name) VALUES ('Alice')");
"#); }

#[test] fn pdo_prepare_execute() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$stmt = $pdo->prepare('SELECT * FROM users WHERE id = ?');
$stmt->execute([1]);
$row = $stmt->fetch();
"#); }

#[test] fn pdo_fetch_all() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$stmt = $pdo->prepare('SELECT * FROM users');
$stmt->execute([]);
$rows = $stmt->fetchAll();
"#); }

#[test] fn pdo_transaction() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("INSERT INTO users (name) VALUES ('Bob')");
$pdo->commit();
"#); }

#[test] fn pdo_rollback() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("DELETE FROM users");
$pdo->rollBack();
"#); }

#[test]
fn pdo_sqlite_memory_runtime_round_trip() {
    let lines = run_prints(r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)');
$stmt = $pdo->prepare('INSERT INTO users (name) VALUES (?)');
$stmt->execute(['Alice']);
$rows = $pdo->query('SELECT name FROM users');
$names = $rows->fetchAll(PDO::FETCH_COLUMN);
echo $names[0];
"#);

    assert_eq!(lines, vec!["Alice"]);
}

// ── mysqli style ────────────────────────────────────────────
#[test] fn mysqli_connect() { compile_ok(r#"<?php $conn = mysqli_connect('sqlite:test.db');"#); }

#[test] fn mysqli_query() { compile_ok(r#"<?php
$conn = mysqli_connect('sqlite:test.db');
$result = mysqli_query($conn, 'SELECT * FROM users');
"#); }

#[test] fn mysqli_close() { compile_ok(r#"<?php
$conn = mysqli_connect('sqlite:test.db');
mysqli_close($conn);
"#); }

#[test] fn mysqli_num_rows() { compile_ok(r#"<?php
$conn = mysqli_connect('sqlite:test.db');
$result = mysqli_query($conn, 'SELECT * FROM users');
echo mysqli_num_rows($result);
"#); }

// ── Real-world patterns ─────────────────────────────────────
#[test] fn crud_pattern() { compile_ok(r#"<?php
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
"#); }

#[test] fn prepared_statements() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:app.db');
$stmt = $pdo->prepare('INSERT INTO users (name, email) VALUES (?, ?)');
$stmt->execute(['Alice', 'alice@example.com']);
$stmt->execute(['Bob', 'bob@example.com']);
"#); }

#[test] fn query_with_loop() { compile_ok(r#"<?php
$pdo = new PDO('sqlite:app.db');
$rows = $pdo->query('SELECT * FROM users');
foreach ($rows as $row) {
    echo $row['name'] . ': ' . $row['email'];
}
"#); }

#[test] fn database_class_wrapper() { compile_ok(r#"<?php
class Database {
    public $pdo;
    public function __construct($dsn) {
        $this->pdo = new PDO($dsn);
    }
    public function query($sql) {
        return $this->pdo->query($sql);
    }
    public function exec($sql) {
        return $this->pdo->exec($sql);
    }
}
$db = new Database('sqlite:app.db');
$db->exec('CREATE TABLE IF NOT EXISTS items (id INTEGER PRIMARY KEY, name TEXT)');
$items = $db->query('SELECT * FROM items');
"#); }
