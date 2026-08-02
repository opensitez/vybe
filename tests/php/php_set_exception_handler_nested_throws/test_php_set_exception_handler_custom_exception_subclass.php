<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_custom_exception_subclass
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

class CustomDomainException extends DomainException {}
$domainCaptured = false;
set_exception_handler(function(Throwable $e) use (&$domainCaptured) {
    if ($e instanceof CustomDomainException) $domainCaptured = true;
});
try {
    throw new CustomDomainException("Domain breach");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $domainCaptured ? "CUSTOM_DOMAIN_EXCEPTION_OK" : "FAIL";
