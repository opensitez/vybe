<?php
// vybe-test: php/type_functions_extended/ctype_alpha_all_letters
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo ctype_alpha('Hello') ? 'yes' : 'no';
echo ctype_alpha('Hello1') ? 'yes' : 'no';
