<?php
// vybe-test: php/database/pdo_execute_bind_value_types
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
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ATTR_ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE metrics (id INTEGER, score REAL, ok INTEGER)');
$stmt = $pdo->prepare('INSERT INTO metrics (id, score, ok) VALUES (:id, :score, :ok)');
$id = 7;
$score = 3.5;
$ok = true;
$stmt->bindValue(':id', $id, PDO::PARAM_INT);
$stmt->bindValue(':score', $score, PDO::PARAM_STR);
$stmt->bindValue(':ok', $ok, PDO::PARAM_BOOL);
$stmt->execute();
$row = $pdo->query('SELECT score, ok FROM metrics')->fetch(PDO::FETCH_NUM);
echo $row[0];
echo '|';
echo ($row[1] === 1 ? 'one' : 'zero');

__vybe_check(ob_get_clean(), "3.5|one");
