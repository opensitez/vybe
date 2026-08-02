<?php
// vybe-test: php/cross_lang/exception_has_cross_lang_shape
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

// PHP throw produces same object shape as Python raise, JS throw, VB Throw
try {
    throw new Exception('something failed');
} catch (Exception $e) {
    // $e has __type, __exception_type, name, message — cross-language compatible
    echo $e;
}
