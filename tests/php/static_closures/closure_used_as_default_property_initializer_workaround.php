<?php
// vybe-test: php/static_closures/closure_used_as_default_property_initializer_workaround
// origin: languages/php/tests/php/test_static_closures.rs

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

class Pipeline {
    private array $stages = [];
    public function pipe(Closure $stage): static {
        $this->stages[] = $stage;
        return $this;
    }
    public function run(mixed $payload): mixed {
        foreach ($this->stages as $stage) {
            $payload = $stage($payload);
        }
        return $payload;
    }
}
$result = (new Pipeline())
    ->pipe(static fn($x) => $x * 2)
    ->pipe(static fn($x) => $x + 1)
    ->run(5);
echo $result;

__vybe_check(ob_get_clean(), "11");
