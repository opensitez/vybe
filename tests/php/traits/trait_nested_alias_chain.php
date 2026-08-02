<?php
// vybe-test: php/traits/trait_nested_alias_chain
// origin: languages/php/tests/php/test_traits.rs

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

trait BaseFlow {
    public function phase(): string { return 'base'; }
}
trait Flow {
    use BaseFlow { phase as basePhase; }
    public function phase(): string { return 'flow'; }
}
class Pipeline {
    use Flow;
    public function run(): string { return $this->phase() . '|' . $this->basePhase(); }
}
echo (new Pipeline())->run();

__vybe_check(ob_get_clean(), "flow|base");
