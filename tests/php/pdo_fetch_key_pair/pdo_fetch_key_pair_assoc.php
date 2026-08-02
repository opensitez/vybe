<?php
// vybe-test: php/pdo_fetch_key_pair/pdo_fetch_key_pair_assoc
// origin: languages/php/tests/php/test_pdo_fetch_key_pair.rs

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

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE settings (key TEXT, value TEXT)");
$pdo->exec("INSERT INTO settings VALUES ('theme', 'dark')");
$pdo->exec("INSERT INTO settings VALUES ('lang', 'en')");

$stmt = $pdo->query("SELECT key, value FROM settings");
$pairs = $stmt->fetchAll(PDO::FETCH_KEY_PAIR);

echo $pairs['theme'] . "|" . $pairs['lang'];

__vybe_check(ob_get_clean(), "dark|en");
