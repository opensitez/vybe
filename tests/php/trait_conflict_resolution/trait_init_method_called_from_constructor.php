<?php
// vybe-test: php/trait_conflict_resolution/trait_init_method_called_from_constructor
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs

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

trait Initializable {
    private bool $initialized = false;
    protected function init(): void { $this->initialized = true; }
    public function isReady(): bool { return $this->initialized; }
}
class Service {
    use Initializable;
    public function __construct() { $this->init(); }
}
echo (new Service())->isReady() ? 'ready' : 'not ready';

__vybe_check(ob_get_clean(), "ready");
