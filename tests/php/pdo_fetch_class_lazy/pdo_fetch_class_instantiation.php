<?php
// vybe-test: php/pdo_fetch_class_lazy/pdo_fetch_class_instantiation
// origin: languages/php/tests/php/test_pdo_fetch_class_lazy.rs

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

class User {
    public string $name;
    public int $age;
    
    public function __construct() {
        // PDO assigns properties BEFORE calling __construct by default
        echo $this->name . ':';
    }
}

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE users (name TEXT, age INTEGER)");
$pdo->exec("INSERT INTO users VALUES ('Alice', 30)");

$stmt = $pdo->query("SELECT name, age FROM users");
$user = $stmt->fetchObject(User::class);
echo $user->age;

__vybe_check(ob_get_clean(), "Alice:30");
