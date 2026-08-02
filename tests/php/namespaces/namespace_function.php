<?php
// vybe-test: php/namespaces/namespace_function
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Utils;
function clamp(int $v, int $lo, int $hi): int {
    return max($lo, min($hi, $v));
}
echo clamp(15, 0, 10);
