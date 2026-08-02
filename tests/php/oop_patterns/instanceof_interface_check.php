<?php
// vybe-test: php/oop_patterns/instanceof_interface_check
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

interface Runnable {
    public function run(): void;
}
interface Stoppable {
    public function stop(): void;
}
class Engine implements Runnable, Stoppable {
    private bool $running = false;
    public function run(): void  { $this->running = true; }
    public function stop(): void { $this->running = false; }
    public function isRunning(): bool { return $this->running; }
}
$e = new Engine();
echo $e instanceof Runnable   ? 'runnable'   : 'not';
echo $e instanceof Stoppable  ? 'stoppable'  : 'not';
