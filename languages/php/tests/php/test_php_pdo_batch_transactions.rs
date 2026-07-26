use super::helpers::run_prints;

fn assert_int(source: &str, expected: i64) {
    assert_eq!(run_prints(source), vec![expected.to_string()]);
}

#[test]
fn php_pdo_positional_batch_insert() {
    for rows in 1..=80_i64 {
        let source = format!(
            r##"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE items (id INTEGER, label TEXT)');
$stmt = $pdo->prepare('INSERT INTO items (id, label) VALUES (?, ?)');
for ($i = 0; $i < {rows}; $i++):
    $stmt->execute([$i, 'value_' . $i]);
endfor;

echo $pdo->query('SELECT COUNT(*) FROM items')->fetchColumn();
"##,
            rows = rows
        );
        assert_int(&source, rows);
    }
}

#[test]
fn php_pdo_named_batch_insert() {
    for rows in 1..=80_i64 {
        let source = format!(
            r##"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE records (id INTEGER, tenant TEXT, flags INTEGER)');
$stmt = $pdo->prepare('INSERT INTO records (id, tenant, flags) VALUES (:id, :tenant, :flags)');
for ($i = 0; $i < {rows}; $i++):
    $stmt->execute([
        ':id' => $i,
        ':tenant' => 'tenant_' . $i,
        ':flags' => $i % 3
    ]);
endfor;

echo $pdo->query('SELECT COUNT(*) FROM records')->fetchColumn();
"##,
            rows = rows
        );
        assert_int(&source, rows);
    }
}

#[test]
fn php_pdo_transaction_commit_rollback() {
    for rows in 1..=60_i64 {
        let commit_source = format!(
            r##"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE tx (id INTEGER, status TEXT)');
$pdo->beginTransaction();
$stmt = $pdo->prepare('INSERT INTO tx (id, status) VALUES (?, ?)');
for ($i = 0; $i < {rows}; $i++):
    $stmt->execute([$i, 'open']);
endfor;
$pdo->commit();
echo $pdo->query('SELECT COUNT(*) FROM tx')->fetchColumn();
"##,
            rows = rows
        );
        assert_int(&commit_source, rows);

        let rollback_source = format!(
            r##"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE tx (id INTEGER, status TEXT)');
$pdo->beginTransaction();
$stmt = $pdo->prepare('INSERT INTO tx (id, status) VALUES (?, ?)');
for ($i = 0; $i < {rows}; $i++):
    $stmt->execute([$i, 'open']);
endfor;
$pdo->rollBack();
$query = $pdo->query('SELECT COUNT(*) FROM tx');
$count = $query->fetchColumn();
echo $count === null ? 0 : $count;
"##,
            rows = rows
        );
        assert_int(&rollback_source, 0);
    }
}

#[test]
fn php_pdo_fetch_shape() {
    for size in 1..=50_i64 {
        let rows = size * 2;
        let source = format!(
            r##"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE data (id INTEGER, tenant TEXT, score INTEGER)');
$pdo->exec(\"INSERT INTO data (id, tenant, score) VALUES (0, 'default', 0), (1, 'default', 10), (2, 'billing', 20), (3, 'billing', 30)\");
$stmt = $pdo->prepare('INSERT INTO data (id, tenant, score) VALUES (?, ?, ?)');
for ($i = 4; $i < {rows}; $i++):
    $tenant = $i % 2 === 0 ? 'billing' : 'default';
    $stmt->execute([$i, $tenant, $i + {rows}]);
endfor;
        $statement = $pdo->prepare('SELECT id, tenant, score FROM data LIMIT ?');
$statement->execute([$size]);
$rows = $statement->fetchAll(PDO::FETCH_ASSOC);
echo count($rows);
"##,
            rows = rows
        );
        assert_int(&source, size);
    }
}
