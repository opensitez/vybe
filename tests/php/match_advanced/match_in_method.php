<?php
// vybe-test: php/match_advanced/match_in_method
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

class Router {
    public function dispatch(string $method, string $path): string {
        return match("$method $path") {
            'GET /'       => 'home',
            'GET /users'  => 'user list',
            'POST /users' => 'create user',
            default       => '404',
        };
    }
}
$r = new Router();
echo $r->dispatch('GET', '/users');
echo $r->dispatch('POST', '/users');
echo $r->dispatch('DELETE', '/anything');
