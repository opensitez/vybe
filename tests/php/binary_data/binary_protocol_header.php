<?php
// vybe-test: php/binary_data/binary_protocol_header
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

// Simulate a simple binary protocol header: magic(2) + version(1) + length(4)
function buildHeader(int $version, int $length): string {
    return pack('nCN', 0xBEEF, $version, $length);
}
function parseHeader(string $data): array {
    return unpack('nmagic/Cversion/Nlength', $data);
}
$header = buildHeader(2, 1024);
$parsed = parseHeader($header);
echo dechex($parsed['magic']) . ':' . $parsed['version'] . ':' . $parsed['length'];
