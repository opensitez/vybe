<?php
// vybe-test: php/compression/compress_cache_value
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

function cacheSerialize(mixed $value): string {
    return base64_encode(gzcompress(serialize($value)));
}
function cacheUnserialize(string $cached): mixed {
    return unserialize(gzuncompress(base64_decode($cached)));
}
$original = ['data' => str_repeat('x', 500), 'count' => 42];
$cached   = cacheSerialize($original);
$restored = cacheUnserialize($cached);
echo $restored['count'] . ':' . strlen($restored['data']);
