<?php
// vybe-test: php/database_prepared/prepared_bindvalue_param_bool_stores_truthy
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
// `on` is a RESERVED WORD in sqlite — unquoted it is a syntax error, and real
// php throws `PDOException: near "on"` on the CREATE. The test is about
// PDO::PARAM_BOOL binding as 1, not about bare keyword identifiers, so quote it.
$pdo->exec('CREATE TABLE f ("on" INTEGER)');
$stmt = $pdo->prepare('INSERT INTO f ("on") VALUES (?)');
$stmt->bindValue(1, true, PDO::PARAM_BOOL);
$stmt->execute();
echo $pdo->query('SELECT "on" FROM f')->fetchColumn();

__vybe_check(ob_get_clean(), "1");
