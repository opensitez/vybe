<?php
// vybe-test: php/binary_data/pack_multiple_chars
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('CCC', 72, 101, 108);
echo $packed;  // "Hel"
