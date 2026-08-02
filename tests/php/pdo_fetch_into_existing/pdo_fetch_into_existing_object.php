<?php
// vybe-test: php/pdo_fetch_into_existing/pdo_fetch_into_existing_object
// origin: languages/php/tests/php/test_pdo_fetch_into_existing.rs

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

class Stats {
    public int $views = 0;
    public int $clicks = 0;
}

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE metrics (views INTEGER, clicks INTEGER)");
$pdo->exec("INSERT INTO metrics VALUES (100, 5)");

$stats = new Stats();
$stmt = $pdo->query("SELECT views, clicks FROM metrics");
$stmt->setFetchMode(PDO::FETCH_INTO, $stats);
$stmt->fetch();

echo $stats->views . "|" . $stats->clicks;

__vybe_check(ob_get_clean(), "100|5");
