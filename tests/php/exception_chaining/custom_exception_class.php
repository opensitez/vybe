<?php
// vybe-test: php/exception_chaining/custom_exception_class
// origin: languages/php/tests/php/test_exception_chaining.rs

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

class DomainException2 extends RuntimeException {
    public function __construct(string $msg, private string $domain) {
        parent::__construct($msg);
    }
    public function getDomain(): string { return $this->domain; }
}
try { throw new DomainException2('err', 'payments'); }
catch (DomainException2 $e) { echo $e->getMessage() . ':' . $e->getDomain(); }

__vybe_check(ob_get_clean(), "err:payments");
