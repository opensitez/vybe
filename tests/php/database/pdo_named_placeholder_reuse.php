<?php
// vybe-test: php/database/pdo_named_placeholder_reuse
// origin: languages/php/tests/php/test_database.rs

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
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE logs (msg TEXT, lvl TEXT)');
$stmt = $pdo->prepare('INSERT INTO logs (msg, lvl) VALUES (:m, :l)');
$stmt->execute([':m' => 'boot', ':l' => 'info']);
$stmt->execute([':m' => 'done', ':l' => 'info']);
echo $pdo->query('SELECT COUNT(*) FROM logs')->fetchColumn();

__vybe_check(ob_get_clean(), "2");
