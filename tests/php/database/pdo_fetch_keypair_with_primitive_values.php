<?php
// vybe-test: php/database/pdo_fetch_keypair_with_primitive_values
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
$pdo->exec('CREATE TABLE kv (k TEXT, v TEXT)');
$pdo->exec("INSERT INTO kv VALUES ('a', 'x'), ('b', 'y')");
$row = $pdo->query('SELECT k, v FROM kv')->fetchAll(PDO::FETCH_KEY_PAIR);
echo is_array($row) ? 'arr' : 'bad';
echo '|';
echo $row['a'];

__vybe_check(ob_get_clean(), "arr|x");
