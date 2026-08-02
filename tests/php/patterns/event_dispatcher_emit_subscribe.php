<?php
// vybe-test: php/patterns/event_dispatcher_emit_subscribe
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

class Dispatcher {
    private $handlers = [];
    public function on(string $event, callable $fn): void { $this->handlers[$event][] = $fn; }
    public function emit(string $event, $payload = null): void {
        foreach ($this->handlers[$event] ?? [] as $fn) { $fn($payload); }
    }
}
$d = new Dispatcher();
$d->on('data', fn($v) => print("got:$v\n"));
$d->on('data', fn($v) => print("also:$v\n"));
$d->emit('data', 42);

__vybe_check(ob_get_clean(), "got:42\nalso:42");
