<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_event_loop_cooperative_scheduler
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

class SimpleLoop {
    private array $fibers = [];
    public function enqueue(Fiber $f): void { $this->fibers[] = $f; }
    public function run(): void {
        while (!empty($this->fibers)) {
            $f = array_shift($this->fibers);
            if (!$f->isStarted()) { $f->start(); }
            elseif ($f->isSuspended()) { $f->resume(); }
            if ($f->isSuspended()) { $this->fibers[] = $f; }
        }
    }
}

$loop = new SimpleLoop();
$loop->enqueue(new Fiber(function() { echo "Task 1\n"; Fiber::suspend(); echo "Task 1 Done\n"; }));
$loop->enqueue(new Fiber(function() { echo "Task 2\n"; Fiber::suspend(); echo "Task 2 Done\n"; }));
$loop->run();
