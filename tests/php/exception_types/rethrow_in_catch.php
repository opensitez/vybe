<?php
// vybe-test: php/exception_types/rethrow_in_catch
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

function process(): void {
    try {
        throw new RuntimeException('original');
    } catch (RuntimeException $e) {
        throw new Exception('wrapped: ' . $e->getMessage(), 0, $e);
    }
}
try { process(); } catch (Exception $e) { echo $e->getMessage(); }
