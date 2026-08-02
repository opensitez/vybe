<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_fetch_key_pair
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE settings (setting_key TEXT PRIMARY KEY, setting_val TEXT)");
$pdo->exec("INSERT INTO settings VALUES ('theme', 'dark'), ('lang', 'en')");

$stmt = $pdo->query("SELECT setting_key, setting_val FROM settings");
$settings = $stmt->fetchAll(PDO::FETCH_KEY_PAIR);
echo "theme=" . $settings["theme"] . " lang=" . $settings["lang"];

__vybe_check(ob_get_clean(), "theme=dark lang=en");
