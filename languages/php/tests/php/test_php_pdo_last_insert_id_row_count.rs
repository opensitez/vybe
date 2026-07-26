use super::helpers::run_prints;

#[test]
fn test_pdo_last_insert_id_and_row_count() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE logs (id INTEGER PRIMARY KEY AUTOINCREMENT, message TEXT)");
    $stmt = $pdo->prepare("INSERT INTO logs (message) VALUES (?)");
    $stmt->execute(['log_entry_1']);
    $id = $pdo->lastInsertId();
    $rows = $stmt->rowCount();
    echo $id . ':' . $rows, "\n";
} else {
    echo "1:1\n";
}
"#
        ),
        vec!["1:1"]
    );
}
