<?php
// vybe-test: php/patterns/mediator_centralized_communication
// origin: languages/php/tests/php/test_patterns.rs

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

class ChatRoom {
    private $log = [];
    public function send(string $from, string $to, string $msg): void {
        $this->log[] = "$from->$to:$msg";
    }
    public function getLog(): array { return $this->log; }
}
class User {
    public function __construct(private string $name, private ChatRoom $room) {}
    public function send(string $to, string $msg): void { $this->room->send($this->name, $to, $msg); }
}
$room = new ChatRoom();
$alice = new User('Alice', $room);
$bob = new User('Bob', $room);
$alice->send('Bob', 'hi');
$bob->send('Alice', 'hello');
foreach ($room->getLog() as $entry) { echo $entry; }

__vybe_check(ob_get_clean(), "Alice->Bob:hiBob->Alice:hello");
