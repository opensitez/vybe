//! Runtime behavior for hash, HMAC, and checksum helpers.

crate::php_cases! {
    hash_sha256_hex_length => {
        r#"<?php
$h = hash('sha256', 'hello world');
echo (strlen($h) === 64 ? 'ok' : 'fail') . (ctype_xdigit($h) ? ':hex' : ':not hex');
"#,
        ["ok:hex"]
    };

    hash_sha512_hex_length => {
        r#"<?php
$h = hash('sha512', 'hello world');
echo (strlen($h) === 128 ? 'ok' : 'fail') . (ctype_xdigit($h) ? ':hex' : ':not hex');
"#,
        ["ok:hex"]
    };

    hash_md5_matches_md5_builtin => {
        r#"<?php
$h = hash('md5', 'hello');
echo (strlen($h) === 32 ? 'ok' : 'fail') . ($h === md5('hello') ? ':matches' : ':differs');
"#,
        ["ok:matches"]
    };

    hash_sha1_matches_sha1_builtin => {
        r#"<?php
$h = hash('sha1', 'hello');
echo (strlen($h) === 40 ? 'ok' : 'fail') . ($h === sha1('hello') ? ':matches' : ':differs');
"#,
        ["ok:matches"]
    };

    hash_hmac_sha256_length => {
        r#"<?php
$mac = hash_hmac('sha256', 'data to sign', 'secret');
echo (strlen($mac) === 64 ? 'ok' : 'fail') . (ctype_xdigit($mac) ? ':hex' : ':not hex');
"#,
        ["ok:hex"]
    };

    crc32_is_int_and_deterministic => {
        r#"<?php
$checksum = crc32('hello world');
echo (is_int($checksum) ? 'int' : 'not int') . (crc32('hello world') === $checksum ? ':deterministic' : ':varies');
"#,
        ["int:deterministic"]
    };

    md5_known_digest => {
        r#"<?php
echo md5('hello');
"#,
        ["5d41402abc4b2a76b9719d911017c592"]
    };

    sha1_matches_builtin_for_hello => {
        r#"<?php
echo sha1('hello');
"#,
        ["aaf4c61ddcc5e8a2dabede0f3b482cd9aea9434d"]
    };
}
