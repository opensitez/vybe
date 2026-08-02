<?php
// vybe-test: php/enums_deep/enum_as_array_key_via_value
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum HttpMethod: string {
    case GET    = 'GET';
    case POST   = 'POST';
    case PUT    = 'PUT';
    case DELETE = 'DELETE';
}
$handlers = [
    HttpMethod::GET->value    => fn() => 'list',
    HttpMethod::POST->value   => fn() => 'create',
    HttpMethod::DELETE->value => fn() => 'delete',
];
$method = HttpMethod::POST;
echo ($handlers[$method->value])();
