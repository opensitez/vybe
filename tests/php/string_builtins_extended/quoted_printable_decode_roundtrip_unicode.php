<?php
// vybe-test: php/string_builtins_extended/quoted_printable_decode_roundtrip_unicode
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$encoded = quoted_printable_encode("Cafe: Crème");
echo quoted_printable_decode($encoded);
