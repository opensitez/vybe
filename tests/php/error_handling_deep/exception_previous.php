<?php
// vybe-test: php/error_handling_deep/exception_previous
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

try {
    try {
        throw new \RuntimeException("original");
    } catch (\RuntimeException $e) {
        throw new \LogicException("wrapped", 0, $e);
    }
} catch (\LogicException $e) {
    echo $e->getMessage();
    echo ':' . $e->getPrevious()->getMessage();
}
