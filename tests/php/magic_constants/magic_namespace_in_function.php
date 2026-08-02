<?php
// vybe-test: php/magic_constants/magic_namespace_in_function
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

namespace Domain\Models;
function currentNs(): string { return __NAMESPACE__; }
echo currentNs();
