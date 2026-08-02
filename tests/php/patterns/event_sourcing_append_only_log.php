<?php
// vybe-test: php/patterns/event_sourcing_append_only_log
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

class EventStore {
    private $events = [];
    public function append(string $type, array $payload): void {
        $this->events[] = ['type' => $type, 'payload' => $payload];
    }
    public function replay(callable $reducer, $initial) {
        return array_reduce($this->events, fn($state, $e) => $reducer($state, $e), $initial);
    }
    public function count(): int { return count($this->events); }
}
$store = new EventStore();
$store->append('deposit', ['amount' => 100]);
$store->append('deposit', ['amount' => 50]);
$store->append('withdraw', ['amount' => 30]);
$balance = $store->replay(function($bal, $e) {
    if ($e['type'] === 'deposit') return $bal + $e['payload']['amount'];
    if ($e['type'] === 'withdraw') return $bal - $e['payload']['amount'];
    return $bal;
}, 0);
echo $balance;
echo $store->count();

__vybe_check(ob_get_clean(), "1203");
