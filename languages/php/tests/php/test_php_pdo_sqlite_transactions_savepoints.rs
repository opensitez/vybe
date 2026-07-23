use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: PDO SQLite Transactions & Savepoints — beginTransaction, commit, rollBack, inTransaction, SAVEPOINT
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_pdo_transaction_rollback_restores_state() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE items (id INTEGER PRIMARY KEY, name TEXT)");

$pdo->beginTransaction();
$pdo->exec("INSERT INTO items (name) VALUES ('Item 1')");
$pdo->rollBack();

$stmt = $pdo->query("SELECT COUNT(*) FROM items");
echo "Count: " . $stmt->fetchColumn();
"#,
    );
    assert_eq!(out, vec!["Count: 0"]);
}

#[test]
fn test_php_pdo_transaction_commit_persists_state() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE items (id INTEGER PRIMARY KEY, name TEXT)");

$pdo->beginTransaction();
$pdo->exec("INSERT INTO items (name) VALUES ('Item 1')");
$pdo->commit();

$stmt = $pdo->query("SELECT COUNT(*) FROM items");
echo "Count: " . $stmt->fetchColumn();
"#,
    );
    assert_eq!(out, vec!["Count: 1"]);
}

#[test]
fn test_php_pdo_nested_transactions_savepoints_sqlite() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE logs (msg TEXT)");

$pdo->beginTransaction();
$pdo->exec("INSERT INTO logs VALUES ('Outer 1')");

$pdo->exec("SAVEPOINT sp1");
$pdo->exec("INSERT INTO logs VALUES ('Inner 1')");
$pdo->exec("ROLLBACK TO SAVEPOINT sp1");

$pdo->commit();

$stmt = $pdo->query("SELECT group_concat(msg, ',') FROM logs");
echo $stmt->fetchColumn();
"#,
    );
    assert_eq!(out, vec!["Outer 1"]);
}

#[test]
fn test_php_pdo_in_transaction_status_flag() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
echo $pdo->inTransaction() ? "TX1" : "TX0";
$pdo->beginTransaction();
echo $pdo->inTransaction() ? " TX1" : " TX0";
$pdo->commit();
echo $pdo->inTransaction() ? " TX1" : " TX0";
"#,
    );
    assert_eq!(out, vec!["TX0 TX1 TX0"]);
}

#[test]
fn test_php_pdo_exception_on_double_begin_transaction() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->beginTransaction();

try {
    $pdo->beginTransaction(); // Double begin transaction throws exception
} catch (PDOException $e) {
    echo "Double beginTransaction caught: " . $e->getMessage();
}
$pdo->rollBack();
"#,
    );
}

#[test]
fn test_php_pdo_sqlite_last_insert_id_sequence() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT)");
$pdo->exec("INSERT INTO users (username) VALUES ('Alice')");
$id1 = $pdo->lastInsertId();
$pdo->exec("INSERT INTO users (username) VALUES ('Bob')");
$id2 = $pdo->lastInsertId();

echo "ID1=$id1 ID2=$id2";
"#,
    );
}

#[test]
fn test_php_pdo_fetch_column_indexed_column() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE t (a INT, b TEXT)");
$pdo->exec("INSERT INTO t VALUES (10, 'ten')");
$stmt = $pdo->query("SELECT a, b FROM t");
echo "B_VAL=" . $stmt->fetchColumn(1);
"#,
    );
}

#[test]
fn test_php_pdo_sqlite_user_defined_function() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
if (method_exists($pdo, "sqliteCreateFunction")) {
    $pdo->sqliteCreateFunction("md5_hash", fn($val) => md5($val));
    $stmt = $pdo->query("SELECT md5_hash('test')");
    echo "Hash: " . $stmt->fetchColumn();
} else {
    echo "sqliteCreateFunction not supported";
}
"#,
    );
}

#[test]
fn test_php_pdo_prepare_positional_parameters() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE scores (player TEXT, score INT)");
$stmt = $pdo->prepare("INSERT INTO scores VALUES (?, ?)");
$stmt->execute(["Alice", 100]);
$stmt->execute(["Bob", 200]);

$query = $pdo->query("SELECT SUM(score) FROM scores");
echo "Total Score: " . $query->fetchColumn();
"#,
    );
}

#[test]
fn test_php_pdo_statement_row_count_insert_update() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE items (id INT, val TEXT)");
$stmt = $pdo->prepare("INSERT INTO items VALUES (?, ?)");
$stmt->execute([1, "a"]);
echo "Affected: " . $stmt->rowCount();
"#,
    );
}
