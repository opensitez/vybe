<?php
// vybe-test: php/spl_autoload/class_alias_reuses_existing_when_preferred
// origin: languages/php/tests/php/test_spl_autoload.rs

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

class Proto {}
class_alias(Proto::class, 'AliasProto');
echo (new AliasProto()) instanceof Proto ? 'yes' : 'no';
echo '|';
echo is_subclass_of(AliasProto::class, Proto::class) ? 'sub' : 'no';

__vybe_check(ob_get_clean(), "yes|no");
