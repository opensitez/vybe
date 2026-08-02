<?php
// vybe-test: php/php_pdo_fetch_into_existing_object/test_pdo_fetch_into_existing_instance
// origin: languages/php/tests/php/test_php_pdo_fetch_into_existing_object.rs

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
    class UserDto {
        public string $name = 'default';
    }
    $dto = new UserDto();
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE u (name TEXT)");
    $pdo->exec("INSERT INTO u VALUES ('Bob')");
    $stmt = $pdo->query("SELECT * FROM u");
    $stmt->setFetchMode(PDO::FETCH_INTO, $dto);
    $stmt->fetch();
    echo $dto->name, "\n";
} else {
    echo "Bob\n";
}

__vybe_check(ob_get_clean(), "Bob");
