<?php
// vybe-test: php/oop_interfaces/interface_default_like_methods_in_runtime_chain
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

interface Escaper {
    public function escape(string $v): string;
}
class HtmlSafe implements Escaper {
    public function escape(string $v): string {
        return str_replace('<', '&lt;', $v);
    }
}
class NoEscape implements Escaper {
    public function escape(string $v): string {
        return $v;
    }
}
function render(Escaper $e, string $v): string {
    return $e->escape($v);
}
echo render(new HtmlSafe, "<x>") . '|' . render(new NoEscape, "<x>");

__vybe_check(ob_get_clean(), "&lt;x>|<x>");
