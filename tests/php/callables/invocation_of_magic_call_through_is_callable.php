<?php
// vybe-test: php/callables/invocation_of_magic_call_through_is_callable
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

class MagicApi {
    private function hidden(string $s): string { return "h:$s"; }
    public function __call(string $name, array $args): string { return "m:$name(" . $args[0] . ")"; }
}
$m = new MagicApi();
echo is_callable([$m, 'dynamic']) ? 'yes' : 'no';
echo '|' . $m->dynamic('x');

__vybe_check(ob_get_clean(), "yes|m:dynamic(x)");
