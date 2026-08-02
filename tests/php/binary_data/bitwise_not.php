<?php
// vybe-test: php/binary_data/bitwise_not
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

echo ~0 & 0xFF;   // 255
echo ~1 & 0xFF;   // 254
