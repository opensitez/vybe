<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_custom_exception_context_payload
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

class HttpPayloadException extends Exception {
    public function __construct(public array $context, string $message = "", int $code = 0) {
        parent::__construct($message, $code);
    }
}

try {
    throw new HttpPayloadException(["ip" => "127.0.0.1"], "Unauthorized", 401);
} catch (HttpPayloadException $e) {
    echo "IP=" . $e->context["ip"] . " Code=" . $e->getCode();
}
