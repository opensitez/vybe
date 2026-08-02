<?php
// vybe-test: php/design_patterns/observer_pattern
// origin: languages/php/tests/php/test_design_patterns.rs

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

interface Observer { public function update(string $event, mixed $data): void; }
class EventEmitter {
    private array $observers = [];
    public function subscribe(string $event, Observer $o): void { $this->observers[$event][] = $o; }
    public function emit(string $event, mixed $data = null): void {
        foreach ($this->observers[$event] ?? [] as $o) $o->update($event, $data);
    }
}
class Logger implements Observer {
    public array $log = [];
    public function update(string $e, mixed $d): void { $this->log[] = "$e:$d"; }
}
$emitter = new EventEmitter;
$logger = new Logger;
$emitter->subscribe('login', $logger);
$emitter->emit('login', 'Alice');
$emitter->emit('login', 'Bob');
echo implode(',', $logger->log);

__vybe_check(ob_get_clean(), "login:Alice,login:Bob");
