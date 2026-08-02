<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_range_exception_value_domain_error
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

function setPercentage(float $pct) {
    if ($pct < 0.0 || $pct > 1.0) {
        throw new RangeException("Percentage must be between 0.0 and 1.0");
    }
}

try {
    setPercentage(1.5);
} catch (RangeException $e) {
    echo $e->getMessage();
}
