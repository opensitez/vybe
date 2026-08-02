<?php
// vybe-test: php/traits/trait_property_defaults_and_override
// origin: languages/php/tests/php/test_traits.rs

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

trait Defaults {
    public function base(): string { return 'd'; }
}
class ParentThing {
    public string $v = 'parent';
}
class ChildThing extends ParentThing {
    use Defaults;
    public string $v = 'child';
}
echo (new ChildThing())->v . (new ChildThing())->base();

__vybe_check(ob_get_clean(), "childd");
