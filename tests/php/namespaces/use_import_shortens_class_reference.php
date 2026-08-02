<?php
// vybe-test: php/namespaces/use_import_shortens_class_reference
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

namespace App\Http {
    class Request {}
}
namespace App\Controllers {
    use App\Http\Request;
    function make(): string {
        return (new Request()) instanceof Request ? 'req' : 'no';
    }
}
echo \App\Controllers\make();

__vybe_check(ob_get_clean(), "req");
