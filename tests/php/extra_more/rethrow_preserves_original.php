<?php
// vybe-test: php/extra_more/rethrow_preserves_original
// origin: languages/php/tests/php/test_extra_more.rs

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

try{
    try{throw new RuntimeException('orig',404);}
    catch(RuntimeException $e){throw new LogicException('wrap',0,$e);}
}catch(LogicException $e){echo $e->getPrevious()->getCode();}

__vybe_check(ob_get_clean(), "404");
