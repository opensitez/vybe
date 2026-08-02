<?php
// vybe-test: php/exception_types/domain_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new DomainException('value outside domain');
} catch (DomainException $e) {
    echo $e->getMessage();
}
