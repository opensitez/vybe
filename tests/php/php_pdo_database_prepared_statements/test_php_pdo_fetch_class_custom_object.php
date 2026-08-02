<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_fetch_class_custom_object
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs
// vybe-test-mode: compile

class UserEntity {
    public int $id;
    public string $name;
    public function getUpperName(): string { return strtoupper($this->name); }
}

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE users (id INT, name TEXT)");
$pdo->exec("INSERT INTO users VALUES (1, 'john')");

$stmt = $pdo->query("SELECT * FROM users");
$user = $stmt->fetchObject(UserEntity::class);
echo $user->getUpperName();
