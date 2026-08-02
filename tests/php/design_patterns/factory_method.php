<?php
// vybe-test: php/design_patterns/factory_method
// origin: languages/php/tests/php/test_design_patterns.rs

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

abstract class Notification {
    abstract public function send(string $msg): string;
    public static function create(string $type): self {
        return match($type) {
            'email' => new EmailNotification,
            'sms'   => new SmsNotification,
            default => throw new InvalidArgumentException("Unknown: $type"),
        };
    }
}
class EmailNotification extends Notification { public function send(string $m): string { return "email:$m"; } }
class SmsNotification extends Notification { public function send(string $m): string { return "sms:$m"; } }
echo Notification::create('email')->send('hello') . ',' . Notification::create('sms')->send('hi');

__vybe_check(ob_get_clean(), "email:hello,sms:hi");
