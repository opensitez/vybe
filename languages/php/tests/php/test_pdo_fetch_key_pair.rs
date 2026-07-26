
crate::php_cases! {
    pdo_fetch_key_pair_assoc => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE settings (key TEXT, value TEXT)");
$pdo->exec("INSERT INTO settings VALUES ('theme', 'dark')");
$pdo->exec("INSERT INTO settings VALUES ('lang', 'en')");

$stmt = $pdo->query("SELECT key, value FROM settings");
$pairs = $stmt->fetchAll(PDO::FETCH_KEY_PAIR);

echo $pairs['theme'] . "|" . $pairs['lang'];
"#,
        ["dark|en"]
    };

    pdo_fetch_key_pair_fails_on_three_columns => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE settings (key TEXT, value TEXT, extra TEXT)");
$pdo->exec("INSERT INTO settings VALUES ('a', '1', 'x')");

$stmt = $pdo->query("SELECT * FROM settings");
try {
    $stmt->fetchAll(PDO::FETCH_KEY_PAIR);
    echo "success";
} catch (\PDOException $e) {
    echo "error";
}
"#,
        ["error"]
    };
}
