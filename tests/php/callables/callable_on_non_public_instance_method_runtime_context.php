<?php
// vybe-test: php/callables/callable_on_non_public_instance_method_runtime_context
// origin: languages/php/tests/php/test_callables.rs

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

class Context {
    private function secret(): string { return 'secret'; }
    public function expose(callable $f): string { return $f($this) . ':' . $this->secret(); }
}
$ctx = new Context();
$f = fn(Context $c): string => 'open';
echo $ctx->expose($f);

__vybe_check(ob_get_clean(), "open:secret");
