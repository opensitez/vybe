<?php
// vybe-test: php/error_handler/user_error_handler_object_method
// origin: languages/php/tests/php/test_error_handler.rs

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

class Logger {
    public array $rows = [];
    public function onError(int $no, string $msg): bool {
        $this->rows[] = $msg;
        return true;
    }
}
$log = new Logger();
set_error_handler([$log, 'onError']);
trigger_error('obj', E_USER_NOTICE);
restore_error_handler();
echo count($log->rows);

__vybe_check(ob_get_clean(), "1");
