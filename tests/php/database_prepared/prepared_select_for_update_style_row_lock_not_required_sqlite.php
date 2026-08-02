<?php
// vybe-test: php/database_prepared/prepared_select_for_update_style_row_lock_not_required_sqlite
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
$pdo->exec('CREATE TABLE bal (id INTEGER, amt INTEGER)');
$pdo->exec('INSERT INTO bal VALUES (1, 100)');
$sel = $pdo->prepare('SELECT amt FROM bal WHERE id = ?');
$sel->execute([1]);
$upd = $pdo->prepare('UPDATE bal SET amt = amt - ? WHERE id = ?');
$upd->execute([30, 1]);
$sel->execute([1]);
echo $sel->fetchColumn();

__vybe_check(ob_get_clean(), "70");
