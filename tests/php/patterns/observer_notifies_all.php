<?php
// vybe-test: php/patterns/observer_notifies_all
// origin: languages/php/tests/php/test_patterns.rs

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

interface Observer {
    public function update(string $event, $data): void;
}
class EventBus {
    private $listeners = [];
    public function subscribe(string $event, Observer $o): void {
        $this->listeners[$event][] = $o;
    }
    public function emit(string $event, $data): void {
        foreach ($this->listeners[$event] ?? [] as $o) {
            $o->update($event, $data);
        }
    }
}
class LogObserver implements Observer {
    private $name;
    public function __construct(string $n) { $this->name = $n; }
    public function update(string $event, $data): void { echo $this->name . ':' . $data; }
}
$bus = new EventBus();
$bus->subscribe('login', new LogObserver('A'));
$bus->subscribe('login', new LogObserver('B'));
$bus->emit('login', 'alice');

__vybe_check(ob_get_clean(), "A:aliceB:alice");
