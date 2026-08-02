<?php
// vybe-test: php/interfaces_deep/interface_implements_and_iterator_pattern_runtime
// origin: languages/php/tests/php/test_interfaces_deep.rs

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

interface Seq {
    public function next(): ?int;
}
class RangeSeq implements Seq {
    private int $current;
    public function __construct(private int $end, int $start = 1) {
        $this->current = $start;
    }
    public function next(): ?int {
        if ($this->current > $this->end) {
            return null;
        }
        return $this->current++;
    }
}
function takeThree(Seq $seq): array {
    $out = [];
    for ($i = 0; $i < 3; $i++) {
        $value = $seq->next();
        if ($value === null) {
            break;
        }
        $out[] = $value;
    }
    return $out;
}
$seq = new RangeSeq(4, 2);
echo implode(',', takeThree($seq));

__vybe_check(ob_get_clean(), "2,3,4");
