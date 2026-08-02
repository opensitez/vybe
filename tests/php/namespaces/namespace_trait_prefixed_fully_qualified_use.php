<?php
// vybe-test: php/namespaces/namespace_trait_prefixed_fully_qualified_use
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
    trait HasToken {
        public function token(): string { return 'tok'; }
    }
}
namespace App {
    class Item {
        use \Core\HasToken;
    }
    echo (new Item())->token();
}

__vybe_check(ob_get_clean(), "tok");
