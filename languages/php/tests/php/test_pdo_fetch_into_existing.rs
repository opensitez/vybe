
crate::php_cases! {
    pdo_fetch_into_existing_object => {
        r#"<?php
class Stats {
    public int $views = 0;
    public int $clicks = 0;
}

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE metrics (views INTEGER, clicks INTEGER)");
$pdo->exec("INSERT INTO metrics VALUES (100, 5)");

$stats = new Stats();
$stmt = $pdo->query("SELECT views, clicks FROM metrics");
$stmt->setFetchMode(PDO::FETCH_INTO, $stats);
$stmt->fetch();

echo $stats->views . "|" . $stats->clicks;
"#,
        ["100|5"]
    };
}
