<?php
// vybe-test: php/oop_interfaces/interface_static_method_called_via_class_implements_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Identifiable {
    public static function label(): string;
}
class Widget implements Identifiable {
    public static function label(): string { return 'widget'; }
}
echo Widget::label();
echo '|';
$ifaces = class_implements(Widget::class);
echo isset($ifaces[Identifiable::class]) ? 'seen' : 'missing';

__vybe_check(ob_get_clean(), "widget|seen");
