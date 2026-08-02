<?php
// vybe-test: php/magic_methods/magic_debuginfo_returns_subset
// origin: languages/php/tests/php/test_magic_methods.rs
// vybe-test-mode: compile

class Config {
    public string $host = "localhost";
    public int $port = 3306;
    private string $dsn = "mysql:host=localhost;port=3306";
    private string $apiKey = "secret-key-12345";
    public function __debugInfo(): array {
        return [
            "host" => $this->host,
            "port" => $this->port,
            "dsn"  => substr($this->dsn, 0, 10) . "...",
        ];
    }
}
$cfg = new Config();
$info = $cfg->__debugInfo();
echo count($info);
echo isset($info["apiKey"]) ? "exposed" : "hidden";
