<?php
// vybe-test: php/namespaces/namespace_class_exists_with_relative_vs_fqcn
// origin: languages/php/tests/php/test_namespaces.rs

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

namespace Core {
    class Engine {}
}
namespace App {
function check(): string {
    $same = class_exists('Engine');
    $fqcn = class_exists('Core\\Engine');
    $withSlash = class_exists('\\Core\\Engine');
        return ($same ? 'same:' : 'same=no:') . ($fqcn ? 'fqcn' : 'nfqcn') . '|' . ($withSlash ? 'slash' : 'nslash');
    }
}
echo \App\check();

__vybe_check(ob_get_clean(), "same=no:fqcn|slash");
