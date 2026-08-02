<?php
// vybe-test: php/components/exception_cross_compat
// origin: languages/php/tests/php/test_components.rs
// vybe-test-mode: compile

try {
    throw new RuntimeException('something broke');
} catch (RuntimeException $e) {
    echo $e;
}
