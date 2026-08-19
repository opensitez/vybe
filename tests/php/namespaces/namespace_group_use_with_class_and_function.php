<?php
// vybe-test: php/namespaces/namespace_group_use_with_class_and_function
// origin: languages/php/tests/php/test_namespaces.rs

namespace {
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
}

namespace Core {
    class Box { public function name(): string { return 'box'; } }
    function helper(string $name): string { return "h:$name"; }
}
namespace App {
    use Core\{Box, function helper};
    $b = new Box();
    echo $b->name();
    echo '|';
    echo helper('x');
}

namespace {
__vybe_check(ob_get_clean(), "box|h:x");
}
