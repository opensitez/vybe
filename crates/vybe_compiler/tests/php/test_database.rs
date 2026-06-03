use super::helpers::{compile_ok, run_prints};

// ── PDO style ───────────────────────────────────────────────
#[test]
fn pdo_connect_sqlite() {
    compile_ok(r#"<?php $pdo = new PDO('sqlite:test.db');"#);
}
#[test]
fn pdo_connect_mysql() {
    compile_ok(r#"<?php $pdo = new PDO('mysql:host=localhost;dbname=mydb');"#);
}
#[test]
fn pdo_connect_postgres() {
    compile_ok(r#"<?php $pdo = new PDO('postgresql://localhost/mydb');"#);
}

#[test]
fn pdo_query() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$rows = $pdo->query('SELECT * FROM users');
"#,
    );
}

#[test]
fn pdo_exec() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->exec('CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT)');
$pdo->exec("INSERT INTO users (name) VALUES ('Alice')");
"#,
    );
}

#[test]
fn pdo_prepare_execute() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$stmt = $pdo->prepare('SELECT * FROM users WHERE id = ?');
$stmt->execute([1]);
$row = $stmt->fetch();
"#,
    );
}

#[test]
fn pdo_fetch_all() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$stmt = $pdo->prepare('SELECT * FROM users');
$stmt->execute([]);
$rows = $stmt->fetchAll();
"#,
    );
}

#[test]
fn pdo_transaction() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("INSERT INTO users (name) VALUES ('Bob')");
$pdo->commit();
"#,
    );
}

#[test]
fn pdo_rollback() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("DELETE FROM users");
$pdo->rollBack();
"#,
    );
}

#[test]
fn pdo_sqlite_memory_runtime_round_trip() {
    let lines = run_prints(
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)');
$stmt = $pdo->prepare('INSERT INTO users (name) VALUES (?)');
$stmt->execute(['Alice']);
$rows = $pdo->query('SELECT name FROM users');
$names = $rows->fetchAll(PDO::FETCH_COLUMN);
echo $names[0];
"#,
    );

    assert_eq!(lines, vec!["Alice"]);
}

#[test]
fn pdo_prepare_select_fetch_assoc_runtime_round_trip() {
    let lines = run_prints(
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)');
$stmt = $pdo->prepare('INSERT INTO users (name) VALUES (?)');
$stmt->execute(['Alice']);
$select = $pdo->prepare('SELECT name FROM users');
$select->execute();
$rows = $select->fetchAll(PDO::FETCH_ASSOC);
echo count($rows);
echo $rows[0]['name'];
"#,
    );

    assert_eq!(lines, vec!["1", "Alice"]);
}

#[test]
fn pdo_bind_param_named_fetch_assoc_runtime_round_trip() {
    let lines = run_prints(
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE projects (id INTEGER PRIMARY KEY, name TEXT)');
$pdo->exec("INSERT INTO projects (id, name) VALUES (10900, 'CTO')");
$stmt = $pdo->prepare('SELECT name FROM projects WHERE id = :id');
$id = 10900;
$stmt->bindParam(':id', $id, PDO::PARAM_INT);
$stmt->execute();
$row = $stmt->fetch(PDO::FETCH_ASSOC);
echo $row['name'];
"#,
    );

    assert_eq!(lines, vec!["CTO"]);
}

// ── mysqli style ────────────────────────────────────────────
#[test]
fn mysqli_connect() {
    compile_ok(r#"<?php $conn = mysqli_connect('sqlite:test.db');"#);
}

#[test]
fn mysqli_query() {
    compile_ok(
        r#"<?php
$conn = mysqli_connect('sqlite:test.db');
$result = mysqli_query($conn, 'SELECT * FROM users');
"#,
    );
}

#[test]
fn mysqli_close() {
    compile_ok(
        r#"<?php
$conn = mysqli_connect('sqlite:test.db');
mysqli_close($conn);
"#,
    );
}

#[test]
fn mysqli_num_rows() {
    compile_ok(
        r#"<?php
$conn = mysqli_connect('sqlite:test.db');
$result = mysqli_query($conn, 'SELECT * FROM users');
echo mysqli_num_rows($result);
"#,
    );
}

#[test]
fn mysqli_surface_compile() {
    compile_ok(
        r#"<?php
$dbh = mysqli_init();
mysqli_select_db($dbh, 'app');
mysqli_set_charset($dbh, 'utf8mb4');
mysqli_ping($dbh);
mysqli_errno($dbh);
mysqli_affected_rows($dbh);
mysqli_insert_id($dbh);
mysqli_num_fields($dbh);
mysqli_fetch_field($dbh);
mysqli_free_result($dbh);
mysqli_more_results($dbh);
mysqli_next_result($dbh);
mysqli_close($dbh);
mysqli_real_escape_string($dbh, 'hello');
mysqli_character_set_name($dbh);
mysqli_get_client_info();
mysqli_get_server_info($dbh);
"#,
    );
}

#[test]
fn mysqli_surface_runtime_shape() {
    let lines = run_prints(
        r#"<?php
$dbh = mysqli_init();
echo mysqli_ping($dbh) ? 'yes' : 'no';
echo mysqli_set_charset($dbh, 'utf8mb4') ? 'yes' : 'no';
echo mysqli_character_set_name($dbh);
echo mysqli_get_client_info();
echo mysqli_get_server_info($dbh);
echo mysqli_real_escape_string($dbh, "O'Reilly\\\\");
echo mysqli_errno($dbh);
echo mysqli_affected_rows($dbh);
echo mysqli_insert_id($dbh);
"#,
    );

    assert_eq!(lines[0], "no");
    assert_eq!(lines[1], "yes");
    assert_eq!(lines[2], "utf8mb4");
    assert!(lines.iter().any(|line| line.contains("mysqlnd")));
    assert!(lines.iter().any(|line| line == "8.0.0"));
    assert!(lines.iter().any(|line| line.contains("O\\'Reilly\\\\")));
    assert!(lines.iter().filter(|line| line.as_str() == "0").count() >= 1);
}

#[test]
fn mysqli_adapter_db_connect_shape_runtime() {
    let lines = run_prints(
        r#"<?php
mysqli_report(0);
$dbh = mysqli_init();
$ok = mysqli_real_connect($dbh, 'localhost', 'user', 'pass', null, null, null, 0);
echo ($ok ? 'yes' : 'no') . ':' . (isset($dbh->connect_errno) ? 'has' : 'missing');
"#,
    );

    assert_eq!(lines, vec!["no:has"]);
}

#[test]
fn mysqli_adapter_exposes_connect_error_helpers() {
    let lines = run_prints(
        r#"<?php
mysqli_report(0);
$dbh = mysqli_init();
mysqli_real_connect($dbh, 'localhost', 'user', 'pass', null, null, null, 0);
echo mysqli_connect_errno();
echo mysqli_connect_error();
echo mysqli_error($dbh);
"#,
    );

    assert_eq!(lines, vec!["1", "Connection failed", "Connection failed"]);
}

// ── Real-world patterns ─────────────────────────────────────
#[test]
fn crud_pattern() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn prepared_statements() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:app.db');
$stmt = $pdo->prepare('INSERT INTO users (name, email) VALUES (?, ?)');
$stmt->execute(['Alice', 'alice@example.com']);
$stmt->execute(['Bob', 'bob@example.com']);
"#,
    );
}

#[test]
fn query_with_loop() {
    compile_ok(
        r#"<?php
$pdo = new PDO('sqlite:app.db');
$rows = $pdo->query('SELECT * FROM users');
foreach ($rows as $row) {
    echo $row['name'] . ': ' . $row['email'];
}
"#,
    );
}

#[test]
fn database_class_wrapper() {
    compile_ok(
        r#"<?php
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
"#,
    );
}
