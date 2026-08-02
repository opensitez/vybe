<?php
// vybe-test: php/try_catch_nested_handlers/three_separate_catch_blocks_same_try_ordered
// origin: languages/php/tests/php/test_try_catch_nested_handlers.rs

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

function probe(int $kind): string {
    try {
        if ($kind === 1) { throw new InvalidArgumentException('a'); }
        if ($kind === 2) { throw new RuntimeException('b'); }
        throw new LogicException('c');
    } catch (InvalidArgumentException $e) { return 'invalid'; }
    catch (RuntimeException $e) { return 'runtime'; }
    catch (LogicException $e) { return 'logic'; }
}
echo probe(1) . probe(2) . probe(3);

__vybe_check(ob_get_clean(), "invalidruntimelogic");
