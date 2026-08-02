<?php
// vybe-test: php/namespaces/use_statement_alias_shortens_runtime_reference
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

namespace Lib\Collections {
    class Bag {
        public function size(): int { return 2; }
    }
}
namespace App {
    use Lib\Collections\Bag as Container;
    function count_items(): int {
        return (new Container())->size();
    }
}
echo \App\count_items();

__vybe_check(ob_get_clean(), "2");
