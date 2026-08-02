<?php
// vybe-test: php/binary_data/binary_string_length
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$bin = "\x00\x01\x02\xFF\xFE";
echo strlen($bin);  // 5 — counts bytes not chars
