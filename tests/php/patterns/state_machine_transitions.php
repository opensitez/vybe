<?php
// vybe-test: php/patterns/state_machine_transitions
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

interface TrafficState {
    public function next(): TrafficState;
    public function color(): string;
}
class Red implements TrafficState {
    public function next(): TrafficState { return new Green(); }
    public function color(): string { return 'red'; }
}
class Green implements TrafficState {
    public function next(): TrafficState { return new Yellow(); }
    public function color(): string { return 'green'; }
}
class Yellow implements TrafficState {
    public function next(): TrafficState { return new Red(); }
    public function color(): string { return 'yellow'; }
}
$state = new Red();
for ($i = 0; $i < 4; $i++) {
    echo $state->color();
    $state = $state->next();
}

__vybe_check(ob_get_clean(), "redgreenyellowred");
