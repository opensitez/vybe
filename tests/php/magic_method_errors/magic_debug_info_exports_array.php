<?php
// vybe-test: php/magic_method_errors/magic_debug_info_exports_array
// origin: languages/php/tests/php/test_magic_method_errors.rs

class DebugMe {
    private int $secret = 5;
    public function __debugInfo(): array { return ['secret' => $this->secret]; }
}
ob_start();
var_dump(new DebugMe());
$out = ob_get_clean();
echo str_contains($out, 'secret') ? 'debug' : 'no';
