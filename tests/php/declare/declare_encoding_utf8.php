<?php
// vybe-test: php/declare/declare_encoding_utf8
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(encoding='UTF-8');
$s = "héllo";
echo mb_strlen($s);
