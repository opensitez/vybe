<?php
// vybe-test: php/exception_types/catch_union_type
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

function risky(bool $flag): void {
    if ($flag) throw new TypeError('type');
    throw new ValueError('value');
}
foreach ([true, false] as $f) {
    try {
        risky($f);
    } catch (TypeError | ValueError $e) {
        echo $e->getMessage();
    }
}
