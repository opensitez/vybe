<?php
// vybe-test: php/patterns/adapter_wraps_incompatible_interface
// origin: languages/php/tests/php/test_patterns.rs

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

class LegacyPrinter {
    public function printText(string $text): void { echo 'Legacy: ' . $text; }
}
interface ModernPrinter {
    public function print(string $text): void;
}
class PrinterAdapter implements ModernPrinter {
    private $legacy;
    public function __construct(LegacyPrinter $l) { $this->legacy = $l; }
    public function print(string $text): void { $this->legacy->printText($text); }
}
function usePrinter(ModernPrinter $p, string $text): void { $p->print($text); }
usePrinter(new PrinterAdapter(new LegacyPrinter()), 'hello');

__vybe_check(ob_get_clean(), "Legacy: hello");
