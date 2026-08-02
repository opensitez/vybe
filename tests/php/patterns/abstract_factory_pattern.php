<?php
// vybe-test: php/patterns/abstract_factory_pattern
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

interface Button { public function render(): string; }
interface Checkbox { public function check(): string; }
class WinButton implements Button { public function render(): string { return 'win-button'; } }
class WinCheckbox implements Checkbox { public function check(): string { return 'win-check'; } }
class MacButton implements Button { public function render(): string { return 'mac-button'; } }
class MacCheckbox implements Checkbox { public function check(): string { return 'mac-check'; } }
interface GUIFactory {
    public function createButton(): Button;
    public function createCheckbox(): Checkbox;
}
class WinFactory implements GUIFactory {
    public function createButton(): Button { return new WinButton(); }
    public function createCheckbox(): Checkbox { return new WinCheckbox(); }
}
class MacFactory implements GUIFactory {
    public function createButton(): Button { return new MacButton(); }
    public function createCheckbox(): Checkbox { return new MacCheckbox(); }
}
function buildUI(GUIFactory $f) {
    echo $f->createButton()->render();
    echo $f->createCheckbox()->check();
}
buildUI(new WinFactory());
buildUI(new MacFactory());

__vybe_check(ob_get_clean(), "win-buttonwin-checkmac-buttonmac-check");
