<?php
// vybe-test: php/oop_patterns/constructor_optional_parameters
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class HttpRequest {
    public function __construct(
        public readonly string $method  = 'GET',
        public readonly string $path    = '/',
        public readonly array  $headers = [],
        public readonly string $body    = ''
    ) {}
}
$get    = new HttpRequest();
$post   = new HttpRequest('POST', '/submit', ['Content-Type' => 'application/json'], '{}');
echo $get->method . ' ' . $get->path;
echo $post->method . ' ' . $post->path;
