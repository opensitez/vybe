<?php
// vybe-test: php/programs/state_machine_traffic_light
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

class TrafficLight {
    private array $states = ['red', 'green', 'yellow'];
    private int $index = 0;
    public function current(): string { return $this->states[$this->index]; }
    public function advance(): void { $this->index = ($this->index + 1) % count($this->states); }
}
$light = new TrafficLight();
for ($i = 0; $i < 6; $i++) {
    echo $light->current() . "\n";
    $light->advance();
}

__vybe_check(ob_get_clean(), "red\ngreen\nyellow\nred\ngreen\nyellow");
