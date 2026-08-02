<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_subject_detach_stops_notifications
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs

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

class Emitter implements SplSubject {
    public SplObjectStorage $obs;
    public int $count = 0;
    public function __construct() { $this->obs = new SplObjectStorage(); }
    public function attach(SplObserver $o): void { $this->obs->attach($o); }
    public function detach(SplObserver $o): void { $this->obs->detach($o); }
    public function notify(): void { foreach ($this->obs as $o) { $o->update($this); } }
}

class Listener implements SplObserver {
    public int $events = 0;
    public function update(SplSubject $s): void { $this->events++; }
}

$emitter = new Emitter();
$l = new Listener();
$emitter->attach($l);
$emitter->notify();

$emitter->detach($l);
$emitter->notify();

echo "Received events: {$l->events}";

__vybe_check(ob_get_clean(), "Received events: 1");
