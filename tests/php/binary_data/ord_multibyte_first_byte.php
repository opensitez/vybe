<?php
// vybe-test: php/binary_data/ord_multibyte_first_byte
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

echo ord('A');    // 65
echo ord('a');    // 97
echo ord(' ');    // 32
echo ord("\x00"); // 0
echo ord("\xFF"); // 255
