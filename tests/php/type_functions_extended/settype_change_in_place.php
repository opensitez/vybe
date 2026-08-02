<?php
// vybe-test: php/type_functions_extended/settype_change_in_place
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$v = '42';
settype($v, 'integer');
echo $v;
