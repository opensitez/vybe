
crate::php_cases! {
    pdo_fetch_func_callback => {
        r#"<?php
$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE numbers (a INTEGER, b INTEGER)");
$pdo->exec("INSERT INTO numbers VALUES (5, 10)");
$pdo->exec("INSERT INTO numbers VALUES (3, 7)");

$stmt = $pdo->query("SELECT a, b FROM numbers");
$results = $stmt->fetchAll(PDO::FETCH_FUNC, function($a, $b) {
    return $a + $b;
});

echo implode(',', $results);
"#,
        ["15,10"]
    };
}
