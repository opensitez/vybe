<?php
// vybe-test: php/namespaces/namespace_nested
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App\Http\Request;
class Handler {
    public function handle(string $method): string {
        return strtoupper($method);
    }
}
$h = new Handler();
echo $h->handle('get');
