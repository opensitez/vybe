<?php
// vybe-test: php/oop_patterns/private_constructor_factory_pattern
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Database {
    private static ?self $instance = null;
    private function __construct(private string $dsn) {}
    public static function connect(string $dsn): self {
        if (self::$instance === null) {
            self::$instance = new self($dsn);
        }
        return self::$instance;
    }
    public function getDsn(): string { return $this->dsn; }
}
$db1 = Database::connect('mysql://localhost/app');
$db2 = Database::connect('ignored-because-already-connected');
echo $db1->getDsn();
echo ($db1 === $db2) ? 'same' : 'different';
