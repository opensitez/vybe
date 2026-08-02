<?php
// vybe-test: php/database_prepared/prepared_while_fetch_accumulates_labels
// origin: languages/php/tests/php/test_database_prepared.rs

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
$pdo->exec('CREATE TABLE x (n TEXT)');
$pdo->exec("INSERT INTO x VALUES ('a'), ('b')");
$stmt = $pdo->prepare('SELECT n FROM x ORDER BY n');
$stmt->execute();
$out = '';
while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) { $out .= $row['n']; }
echo $out;

__vybe_check(ob_get_clean(), "ab");
