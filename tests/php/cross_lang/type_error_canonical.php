<?php
// vybe-test: php/cross_lang/type_error_canonical
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

try {
    throw new TypeError('wrong type');
} catch (TypeError $e) {
    // Maps to canonical "TypeError" — catchable in JS as TypeError
    echo $e;
}
