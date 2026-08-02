<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_hash_algos_list_availability
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs
// vybe-test-mode: compile

$algos = hash_algos();
echo in_array("sha256", $algos) && in_array("md5", $algos) ? "ALGOS_OK" : "FAIL";
