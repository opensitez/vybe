<?php
// vybe-test: php/binary_data/binary_uuid_from_bytes
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

function bytesToUuid(string $bytes): string {
    $hex = bin2hex($bytes);
    return vsprintf('%s%s-%s-%s-%s-%s%s%s', str_split($hex, 4));
}
$bytes = random_bytes(16);
$uuid = bytesToUuid($bytes);
echo strlen($uuid) === 36 ? 'valid uuid length' : 'wrong length';
echo substr_count($uuid, '-') === 4 ? ':four dashes' : ':wrong dashes';
