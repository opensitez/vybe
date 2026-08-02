<?php
// vybe-test: php/oop_patterns/interface_constant_implementation
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

interface HasVersion {
    const API_VERSION = '2.0';
}
class Client implements HasVersion {
    public function version(): string { return self::API_VERSION; }
}
$c = new Client();
echo $c->version();
echo Client::API_VERSION;
