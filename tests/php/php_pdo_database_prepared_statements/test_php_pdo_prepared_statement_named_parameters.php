<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_prepared_statement_named_parameters
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs

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

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE products (id INTEGER, price REAL)");
$pdo->exec("INSERT INTO products VALUES (1, 19.99), (2, 49.50)");

$stmt = $pdo->prepare("SELECT price FROM products WHERE id = :id");
$stmt->bindValue(":id", 2, PDO::PARAM_INT);
$stmt->execute();
$row = $stmt->fetch(PDO::FETCH_ASSOC);
echo "Price: " . $row["price"];

__vybe_check(ob_get_clean(), "Price: 49.5");
