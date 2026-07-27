crate::php_cases! {
    pdo_fetch_group_by_first_column => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE colors (group_id INTEGER, color TEXT)");
$pdo->exec("INSERT INTO colors VALUES (1, 'red')");
$pdo->exec("INSERT INTO colors VALUES (1, 'blue')");
$pdo->exec("INSERT INTO colors VALUES (2, 'green')");

$stmt = $pdo->query("SELECT group_id, color FROM colors");
$grouped = $stmt->fetchAll(PDO::FETCH_COLUMN | PDO::FETCH_GROUP);

echo count($grouped[1]) . "|" . $grouped[1][0] . "|" . $grouped[2][0];
"#,
        ["2|red|green"]
    };
}
