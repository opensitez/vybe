<?php
// vybe-test: php/interfaces_deep/interface_static_method_dispatch_inheritance_runtime
// origin: languages/php/tests/php/test_interfaces_deep.rs

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
    public static function kind(): string;
}
class Base implements Identifiable {
    public static function kind(): string {
        return 'base';
    }
}
class Child extends Base {
    public static function kind(): string {
        return 'child';
    }
}
function typeKind(string $name): string {
    return $name::kind();
}
echo typeKind(Base::class) . '|' . typeKind(Child::class);

__vybe_check(ob_get_clean(), "base|child");
