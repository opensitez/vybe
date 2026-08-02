<?php
// vybe-test: php/namespaces/namespace_nested_alias_stack
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

namespace Infra {
    class Handler { public function id(): string { return 'handler'; } }
}
namespace App\Services {
    use Infra\Handler as InfraHandler;
    class Facade {
        public function make(): string {
            return (new InfraHandler())->id();
        }
    }
}
namespace App {
    use App\Services\Facade;
    echo (new Facade())->make();
}

__vybe_check(ob_get_clean(), "handler");
