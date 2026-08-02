<?php
// vybe-test: php/weak_references_runtime/weak_map_event_emitter_callback
// origin: languages/php/tests/php/test_weak_references_runtime.rs

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

class EventEmitter {
    private WeakMap $listeners;
    public function __construct() { $this->listeners = new WeakMap(); }
    public function on(object $target, callable $cb): void { $this->listeners[$target] = $cb; }
    public function emit(object $target, string $event): void {
        $cb = $this->listeners[$target] ?? null;
        if ($cb) $cb($event);
    }
}
$emitter = new EventEmitter();
$btn = new stdClass();
$emitter->on($btn, fn($e) => print("Button: $e"));
$emitter->emit($btn, 'click');

__vybe_check(ob_get_clean(), "Button: click");
