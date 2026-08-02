<?php
// vybe-test: php/type_functions_extended/ctype_space_whitespace
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_space("  \t\n") ? 'yes' : 'no';
echo ctype_space('  x  ') ? 'yes' : 'no';
