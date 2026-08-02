<?php
// vybe-test: php/compression/gzuncompress_roundtrip
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$texts = ["", "a", str_repeat("x", 1000), "Unicode: héllo wörld"];
foreach ($texts as $t) {
    $c = gzcompress($t);
    echo gzuncompress($c) === $t ? 'ok ' : 'fail ';
}
