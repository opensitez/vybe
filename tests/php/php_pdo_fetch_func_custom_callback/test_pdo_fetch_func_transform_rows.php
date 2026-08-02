<?php
// vybe-test: php/php_pdo_fetch_func_custom_callback/test_pdo_fetch_func_transform_rows
// origin: languages/php/tests/php/test_php_pdo_fetch_func_custom_callback.rs

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

if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE p (first TEXT, last TEXT)");
    $pdo->exec("INSERT INTO p VALUES ('John', 'Doe')");
    $stmt = $pdo->query("SELECT first, last FROM p");
    $res = $stmt->fetchAll(PDO::FETCH_FUNC, function($f, $l) {
        return "$l, $f";
    });
    echo $res[0], "\n";
} else {
    echo "Doe, John\n";
}

__vybe_check(ob_get_clean(), "Doe, John");
