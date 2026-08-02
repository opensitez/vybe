<?php
// vybe-test: php/database/pdo_fetch_class_into_custom_class
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

class User { public int $id; public string $name; }
$pdo = new PDO('sqlite::memory:');
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE users (id INTEGER, name TEXT)');
$pdo->exec("INSERT INTO users VALUES (5, 'bob')");
$user = $pdo->query('SELECT * FROM users')->fetchObject(User::class);
echo $user->id . ':' . $user->name;

__vybe_check(ob_get_clean(), "5:bob");
