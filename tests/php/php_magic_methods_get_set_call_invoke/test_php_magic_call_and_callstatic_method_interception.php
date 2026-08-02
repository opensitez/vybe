<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_call_and_callstatic_method_interception
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs

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

class MagicProxy {
    public function __call(string $name, array $args): string {
        return "DYNAMIC_$name(" . implode(",", $args) . ")";
    }
    public static function __callStatic(string $name, array $args): string {
        return "STATIC_$name(" . implode(",", $args) . ")";
    }
}

$p = new MagicProxy();
echo $p->findUser(42) . " | " . MagicProxy::whereStatus("active");

__vybe_check(ob_get_clean(), "DYNAMIC_findUser(42) | STATIC_whereStatus(active)");
