<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_crypt_des_md5_sha512_hashes
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs
// vybe-test-mode: compile

$hashed = crypt("my_password", '$6$rounds=5000$usesomesalt$');
echo strlen($hashed) > 0 ? "CRYPT_OK" : "FAIL";
