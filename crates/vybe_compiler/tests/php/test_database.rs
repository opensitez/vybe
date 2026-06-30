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

    assert_eq!(lines, vec!["1Alice"]);
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

    assert_eq!(lines.len(), 1);
    let s = &lines[0];
    assert!(s.starts_with("no"));
    assert!(s.contains("yes"));
    assert!(s.contains("utf8mb4"));
    assert!(s.contains("mysqlnd"));
    assert!(s.contains("8.0.0"));
    assert!(s.contains("O\\'Reilly\\\\"));
    assert!(s.ends_with("0"));
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

// ── PDO runtime (`php_cases!`) ──────────────────────────────────

crate::php_cases! {
    pdo_sqlite_memory_insert_and_last_insert_id => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE posts (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT)');
$pdo->exec("INSERT INTO posts (title) VALUES ('first')");
echo $pdo->lastInsertId();
"#,
        ["1"]
    };

    pdo_prepared_statement_row_count_after_update => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE posts (id INTEGER PRIMARY KEY, title TEXT)');
$pdo->exec("INSERT INTO posts (title) VALUES ('old')");
$stmt = $pdo->prepare("UPDATE posts SET title = 'new' WHERE id = 1");
$stmt->execute();
echo $stmt->rowCount();
"#,
        ["1"]
    };

    pdo_fetch_object_returns_std_class_row => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER, name TEXT)');
$pdo->exec("INSERT INTO users VALUES (1, 'ada')");
$row = $pdo->query('SELECT * FROM users')->fetch(PDO::FETCH_OBJ);
echo $row->name;
"#,
        ["ada"]
    };

    pdo_fetch_class_into_custom_class => {
        r#"<?php
class User { public int $id; public string $name; }
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER, name TEXT)');
$pdo->exec("INSERT INTO users VALUES (5, 'bob')");
$user = $pdo->query('SELECT * FROM users')->fetchObject(User::class);
echo $user->id . ':' . $user->name;
"#,
        ["5:bob"]
    };

    pdo_fetch_column_first_scalar => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE cfg (v TEXT)');
$pdo->exec("INSERT INTO cfg VALUES ('on')");
echo $pdo->query('SELECT v FROM cfg')->fetchColumn();
"#,
        ["on"]
    };

    pdo_transaction_commit_persists_rows => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (id INTEGER, v TEXT)');
$pdo->beginTransaction();
$pdo->exec("INSERT INTO t VALUES (1, 'a')");
$pdo->commit();
echo $pdo->query('SELECT COUNT(*) FROM t')->fetchColumn();
"#,
        ["1"]
    };

    pdo_transaction_rollback_discards_rows => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (id INTEGER, v TEXT)');
$pdo->beginTransaction();
$pdo->exec("INSERT INTO t VALUES (1, 'gone')");
$pdo->rollBack();
echo $pdo->query('SELECT COUNT(*) FROM t')->fetchColumn();
"#,
        ["0"]
    };

    pdo_in_clause_with_positional_placeholders => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE tags (name TEXT)');
$pdo->exec("INSERT INTO tags VALUES ('a'), ('b'), ('c')");
$stmt = $pdo->prepare('SELECT name FROM tags WHERE name IN (?, ?)');
$stmt->execute(['a', 'c']);
echo implode(',', $stmt->fetchAll(PDO::FETCH_COLUMN));
"#,
        ["a,c"]
    };

    pdo_named_placeholder_reuse => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE logs (msg TEXT, lvl TEXT)');
$stmt = $pdo->prepare('INSERT INTO logs (msg, lvl) VALUES (:m, :l)');
$stmt->execute([':m' => 'boot', ':l' => 'info']);
$stmt->execute([':m' => 'done', ':l' => 'info']);
echo $pdo->query('SELECT COUNT(*) FROM logs')->fetchColumn();
"#,
        ["2"]
    };

    pdo_quote_escapes_string_literal => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$q = $pdo->quote("O'Reilly");
echo str_contains($q, "''") || str_contains($q, "\\'") ? 'quoted' : $q;
"#,
        ["quoted"]
    };

    pdo_attr_default_fetch_mode_assoc => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
$pdo->exec('CREATE TABLE x (k TEXT)');
$pdo->exec("INSERT INTO x VALUES ('v')");
echo $pdo->query('SELECT k FROM x')->fetch()['k'];
"#,
        ["v"]
    };

    pdo_exec_returns_affected_row_count => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (n INTEGER)');
echo $pdo->exec('INSERT INTO t (n) VALUES (1), (2), (3)');
"#,
        ["3"]
    };

    pdo_fetch_all_grouped_by_column => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE items (cat TEXT, name TEXT)');
$pdo->exec("INSERT INTO items VALUES ('a', '1'), ('a', '2'), ('b', '3')");
$rows = $pdo->query('SELECT cat, name FROM items')->fetchAll(PDO::FETCH_GROUP | PDO::FETCH_ASSOC);
echo count($rows['a']);
"#,
        ["2"]
    };

    pdo_null_param_binds_as_null => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE t (v TEXT)');
$stmt = $pdo->prepare('INSERT INTO t (v) VALUES (?)');
$stmt->execute([null]);
$row = $pdo->query('SELECT v FROM t')->fetch(PDO::FETCH_ASSOC);
echo $row['v'] === null ? 'null' : 'set';
"#,
        ["null"]
    };

    pdo_get_attribute_driver_name_sqlite => {
        r#"<?php
echo (new PDO('sqlite::memory:'))->getAttribute(PDO::ATTR_DRIVER_NAME);
"#,
        ["sqlite"]
    };

    pdo_prepare_bad_sql_throws_pdo_exception => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
try {
    $pdo->prepare('SELEC bad');
    echo 'ok';
} catch (PDOException $e) {
    echo 'pdo';
}
"#,
        ["pdo"]
    };
}
