<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_abstract_method_dispatch_runtime
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs

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

abstract class Processor {
    abstract protected function transform(string $input): string;
    public function run(string $input): string {
        return $this->transform($input) . '!';
    }
}

class UpperProcessor extends Processor {
    protected function transform(string $input): string {
        return strtoupper($input);
    }
}

echo (new UpperProcessor())->run('ok');

__vybe_check(ob_get_clean(), "OK!");
