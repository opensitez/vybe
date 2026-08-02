<?php
// vybe-test: php/database/pdo_in_clause_with_positional_placeholders
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
$pdo->exec('CREATE TABLE tags (name TEXT)');
$pdo->exec("INSERT INTO tags VALUES ('a'), ('b'), ('c')");
$stmt = $pdo->prepare('SELECT name FROM tags WHERE name IN (?, ?)');
$stmt->execute(['a', 'c']);
echo implode(',', $stmt->fetchAll(PDO::FETCH_COLUMN));

__vybe_check(ob_get_clean(), "a,c");
