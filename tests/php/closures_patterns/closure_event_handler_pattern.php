<?php
// vybe-test: php/closures_patterns/closure_event_handler_pattern
// origin: languages/php/tests/php/test_closures_patterns.rs

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

class Button {
    private array $handlers = [];
    public function onClick(Closure $fn): void { $this->handlers[] = $fn; }
    public function click(string $data): void { foreach ($this->handlers as $h) $h($data); }
}
$log = [];
$btn = new Button;
$btn->onClick(function($d) use (&$log) { $log[] = "clicked:$d"; });
$btn->onClick(function($d) use (&$log) { $log[] = "handled:$d"; });
$btn->click('left');
echo implode(',', $log);

__vybe_check(ob_get_clean(), "clicked:left,handled:left");
