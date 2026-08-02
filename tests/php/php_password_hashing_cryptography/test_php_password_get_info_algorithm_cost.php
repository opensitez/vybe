<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_password_get_info_algorithm_cost
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs
// vybe-test-mode: compile

$hash = password_hash("pwd", PASSWORD_BCRYPT, ["cost" => 5]);
$info = password_get_info($hash);
echo "Algo=" . $info["algoName"] . " Cost=" . $info["options"]["cost"];
