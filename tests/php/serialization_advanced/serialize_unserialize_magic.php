<?php
// vybe-test: php/serialization_advanced/serialize_unserialize_magic
// origin: languages/php/tests/php/test_serialization_advanced.rs

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

class DateRange {
    public function __construct(
        private \DateTimeImmutable $start,
        private \DateTimeImmutable $end
    ) {}
    public function __serialize(): array {
        return ['start' => $this->start->format('Y-m-d'), 'end' => $this->end->format('Y-m-d')];
    }
    public function __unserialize(array $data): void {
        $this->start = new \DateTimeImmutable($data['start']);
        $this->end   = new \DateTimeImmutable($data['end']);
    }
    public function days(): int {
        return (int)$this->start->diff($this->end)->days;
    }
}
$range = new DateRange(new \DateTimeImmutable('2024-01-01'), new \DateTimeImmutable('2024-01-31'));
$s = serialize($range);
$r2 = unserialize($s);
echo $r2->days();

__vybe_check(ob_get_clean(), "30");
