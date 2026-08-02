<?php
// vybe-test: php/interfaces_deep/interface_constant_public_only
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Configurable {
    const string DEFAULT_HOST = 'localhost';
    const int    DEFAULT_PORT = 8080;
}
class Server implements Configurable {
    public string $host;
    public int    $port;
    public function __construct() {
        $this->host = self::DEFAULT_HOST;
        $this->port = self::DEFAULT_PORT;
    }
}
$s = new Server();
echo $s->host . ':' . $s->port;
