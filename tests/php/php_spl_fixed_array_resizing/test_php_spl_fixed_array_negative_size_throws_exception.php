<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_negative_size_throws_exception
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs
// vybe-test-mode: compile

try {
    $fixed = new SplFixedArray(-1);
} catch (InvalidArgumentException | ValueError $e) {
    echo "NEGATIVE_SIZE_EX";
}
