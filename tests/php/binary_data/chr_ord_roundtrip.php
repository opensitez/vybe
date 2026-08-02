<?php
// vybe-test: php/binary_data/chr_ord_roundtrip
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

for ($i = 0; $i < 128; $i++) {
    if (ord(chr($i)) !== $i) { echo "fail at $i"; break; }
}
echo 'all ok';
