<?php
// vybe-test: php/interfaces_deep/abstract_class_partial_interface
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Lifecycle {
    public function start(): void;
    public function stop(): void;
    public function isRunning(): bool;
}
abstract class BaseService implements Lifecycle {
    protected bool $running = false;
    public function isRunning(): bool { return $this->running; }
    // start() and stop() left to subclasses
}
class HttpService extends BaseService {
    public function start(): void { $this->running = true; }
    public function stop(): void  { $this->running = false; }
}
$svc = new HttpService();
echo $svc->isRunning() ? 'running' : 'stopped';
$svc->start();
echo $svc->isRunning() ? ':running' : ':stopped';
