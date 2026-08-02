<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_dynamic_method_name_precedence
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

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

class Service {
    public ?Tag $ghost;
    public function tag(): ?Tag {
        return new Tag();
    }
    public function fallback(): string {
        return 'fb';
    }
    public function __construct() {
        $this->ghost = null;
    }
}
class Tag {
    public function name(): string { return 'ok'; }
}

$service = new Service();
echo $service->tag()?->name() . '|';
echo $service->ghost?->name() ?? 'missing';

__vybe_check(ob_get_clean(), "ok|missing");
