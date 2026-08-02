<?php
// vybe-test: php/serialization_advanced/sleep_wakeup_basic
// origin: languages/php/tests/php/test_serialization_advanced.rs

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

class Cached {
    public string $data;
    private bool $loaded = false;
    public function __construct(string $data) { $this->data = $data; $this->loaded = true; }
    public function __sleep(): array { return ['data']; }
    public function __wakeup(): void { $this->loaded = true; }
    public function isLoaded(): bool { return $this->loaded; }
}
$c = new Cached("important");
$s = serialize($c);
$c2 = unserialize($s);
echo $c2->data . ':' . ($c2->isLoaded() ? 'ready' : 'not');

__vybe_check(ob_get_clean(), "important:ready");
