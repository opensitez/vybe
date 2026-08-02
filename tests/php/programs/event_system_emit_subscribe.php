<?php
// vybe-test: php/programs/event_system_emit_subscribe
// origin: languages/php/tests/php/test_programs.rs

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

class Events {
    private array $listeners = [];
    public function on(string $event, callable $fn): void { $this->listeners[$event][] = $fn; }
    public function emit(string $event, ...$args): void {
        foreach ($this->listeners[$event] ?? [] as $fn) $fn(...$args);
    }
}
$events = new Events();
$log = [];
$events->on('tick', function(int $n) use (&$log) { $log[] = "tick:$n"; });
$events->on('tick', function(int $n) use (&$log) { $log[] = "tock:$n"; });
$events->emit('tick', 1);
$events->emit('tick', 2);
foreach ($log as $entry) echo $entry . "\n";

__vybe_check(ob_get_clean(), "tick:1\ntock:1\ntick:2\ntock:2");
