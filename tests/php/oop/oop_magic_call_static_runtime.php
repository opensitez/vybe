<?php
// vybe-test: php/oop/oop_magic_call_static_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class CommandBus {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
    public static function __callStatic(string $name, array $args): string {
        return strtoupper($name) . ':' . implode('|', $args);
    }
}
$obj = new CommandBus();
echo $obj->render(1, 2);
echo '|';
echo CommandBus::dispatch('job');

__vybe_check(ob_get_clean(), "render:1,2|DISPATCH:job");
