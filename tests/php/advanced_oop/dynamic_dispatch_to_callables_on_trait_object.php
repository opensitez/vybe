<?php
// vybe-test: php/advanced_oop/dynamic_dispatch_to_callables_on_trait_object
// origin: languages/php/tests/php/test_advanced_oop.rs

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

interface Executor {
    public function execute(string $task): string;
}
class SyncExecutor implements Executor {
    public function execute(string $task): string { return "sync:$task"; }
}
class AsyncLikeExecutor implements Executor {
    public function execute(string $task): string { return "async:$task"; }
}
$executors = [new SyncExecutor(), new AsyncLikeExecutor()];
echo $executors[0]->execute('build') . '|' . $executors[1]->execute('build');

__vybe_check(ob_get_clean(), "sync:build|async:build");
