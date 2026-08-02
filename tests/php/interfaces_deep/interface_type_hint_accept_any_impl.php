<?php
// vybe-test: php/interfaces_deep/interface_type_hint_accept_any_impl
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Logger { public function log(string $msg): void; }
class ConsoleLogger implements Logger {
    private array $log = [];
    public function log(string $msg): void { $this->log[] = $msg; }
    public function getLog(): array { return $this->log; }
}
class NullLogger implements Logger { public function log(string $msg): void {} }
function doWork(Logger $logger): void {
    $logger->log('started');
    $logger->log('done');
}
$c = new ConsoleLogger();
doWork($c);
echo count($c->getLog());
doWork(new NullLogger());
echo 'ok';
