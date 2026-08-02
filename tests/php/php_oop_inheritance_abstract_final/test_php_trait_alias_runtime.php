<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_trait_alias_runtime
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs

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

trait First {
    public function kind(): string { return 'first'; }
}
trait Second {
    public function kind(): string { return 'second'; }
}
class Item {
    use First, Second {
        First::kind insteadof Second;
        Second::kind as secondaryKind;
    }
}
$item = new Item();
echo $item->kind();
echo '|';
echo $item->secondaryKind();

__vybe_check(ob_get_clean(), "first|second");
