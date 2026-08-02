<?php
// vybe-test: php/oop/sleep_and_wakeup_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Persisted {
    public string $token;
    public function __construct(string $token) { $this->token = $token; }
    public function __sleep(): array { return ['token']; }
    public function __wakeup(): void { $this->token = $this->token . '-awake'; }
}
$p = new Persisted('abc');
$restored = unserialize(serialize($p));
echo $restored->token;

__vybe_check(ob_get_clean(), "abc-awake");
