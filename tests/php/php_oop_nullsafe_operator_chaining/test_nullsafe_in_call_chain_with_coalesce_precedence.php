<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_in_call_chain_with_coalesce_precedence
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

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

class Level {
    public function name(): string { return ''; }
}
class Holder {
    public ?Level $level = null;
}
class Root {
    public function holder(): ?Holder { return null; }
    public function fallback(): string { return 'fb'; }
}
class RootWithLevel extends Root {
    public Holder $holderObj;
    public function __construct() {
        $this->holderObj = new Holder();
        $this->holderObj->level = new Level();
    }
    public function holder(): ?Holder { return $this->holderObj; }
}

echo (new Root())->holder()?->name() ?? 'no-holder';
echo '|';
echo (new RootWithLevel())->holder()?->level?->name() ?? 'no-name';
echo '|';
echo (new RootWithLevel())->holder()?->level?->name() ?: 'fallback-name';

__vybe_check(ob_get_clean(), "no-holder||fallback-name");
