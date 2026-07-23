use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: PDO & Prepared Statements — PDO, PDOStatement, fetchMode, bindValue, bindParam, transactions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_pdo_in_memory_sqlite_connection() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)");
$pdo->exec("INSERT INTO users (name) VALUES ('Alice'), ('Bob')");

$stmt = $pdo->query("SELECT name FROM users ORDER BY id");
$names = $stmt->fetchAll(PDO::FETCH_COLUMN);
echo implode(", ", $names);
"#,
    );
    assert_eq!(out, vec!["Alice, Bob"]);
}

#[test]
fn test_php_pdo_prepared_statement_named_parameters() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE products (id INTEGER, price REAL)");
$pdo->exec("INSERT INTO products VALUES (1, 19.99), (2, 49.50)");

$stmt = $pdo->prepare("SELECT price FROM products WHERE id = :id");
$stmt->bindValue(":id", 2, PDO::PARAM_INT);
$stmt->execute();
$row = $stmt->fetch(PDO::FETCH_ASSOC);
echo "Price: " . $row["price"];
"#,
    );
    assert_eq!(out, vec!["Price: 49.5"]);
}

#[test]
fn test_php_pdo_fetch_obj_stdclass() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE config (k TEXT, v TEXT)");
$pdo->exec("INSERT INTO config VALUES ('app_env', 'production')");

$stmt = $pdo->query("SELECT k, v FROM config");
$obj = $stmt->fetch(PDO::FETCH_OBJ);
echo "{$obj->k}={$obj->v}";
"#,
    );
    assert_eq!(out, vec!["app_env=production"]);
}

#[test]
fn test_php_pdo_fetch_key_pair() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE settings (setting_key TEXT PRIMARY KEY, setting_val TEXT)");
$pdo->exec("INSERT INTO settings VALUES ('theme', 'dark'), ('lang', 'en')");

$stmt = $pdo->query("SELECT setting_key, setting_val FROM settings");
$settings = $stmt->fetchAll(PDO::FETCH_KEY_PAIR);
echo "theme=" . $settings["theme"] . " lang=" . $settings["lang"];
"#,
    );
    assert_eq!(out, vec!["theme=dark lang=en"]);
}

#[test]
fn test_php_pdo_transaction_commit_and_rollback() {
    let out = run_prints(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE accounts (id INT, balance REAL)");
$pdo->exec("INSERT INTO accounts VALUES (1, 100.0)");

try {
    $pdo->beginTransaction();
    $pdo->exec("UPDATE accounts SET balance = balance - 50.0 WHERE id = 1");
    $pdo->rollBack();
} catch (Exception $e) {}

$stmt = $pdo->query("SELECT balance FROM accounts WHERE id = 1");
echo "Balance: " . $stmt->fetchColumn();
"#,
    );
    assert_eq!(out, vec!["Balance: 100"]);
}

#[test]
fn test_php_pdo_fetch_class_custom_object() {
    compile_ok(
        r#"<?php
class UserEntity {
    public int $id;
    public string $name;
    public function getUpperName(): string { return strtoupper($this->name); }
}

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE users (id INT, name TEXT)");
$pdo->exec("INSERT INTO users VALUES (1, 'john')");

$stmt = $pdo->query("SELECT * FROM users");
$user = $stmt->fetchObject(UserEntity::class);
echo $user->getUpperName();
"#,
    );
}

#[test]
fn test_php_pdo_last_insert_id() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE logs (id INTEGER PRIMARY KEY AUTOINCREMENT, msg TEXT)");
$pdo->exec("INSERT INTO logs (msg) VALUES ('log1')");
echo "ID: " . $pdo->lastInsertId();
"#,
    );
}

#[test]
fn test_php_pdo_bind_param_reference_binding() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE data (val INT)");
$stmt = $pdo->prepare("INSERT INTO data VALUES (?)");

$val = 0;
$stmt->bindParam(1, $val);
$val = 10; $stmt->execute();
$val = 20; $stmt->execute();

$stmt2 = $pdo->query("SELECT SUM(val) FROM data");
echo $stmt2->fetchColumn();
"#,
    );
}

#[test]
fn test_php_pdo_error_code_and_info() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
@$pdo->query("SELECT * FROM non_existent_table");
echo "ErrCode: " . $pdo->errorCode();
print_r($pdo->errorInfo());
"#,
    );
}

#[test]
fn test_php_pdo_sqlite_create_function() {
    compile_ok(
        r#"<?php
$pdo = new PDO("sqlite::memory:");
$pdo->sqliteCreateFunction("php_md5", fn($str) => md5($str));

$stmt = $pdo->query("SELECT php_md5('test')");
echo $stmt->fetchColumn();
"#,
    );
}
