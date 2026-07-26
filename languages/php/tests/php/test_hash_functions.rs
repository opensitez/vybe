//! `md5`, `sha1`, `hash`, `hash_hmac`, and `crc32`.

crate::php_cases! {
    md5_string_produces_32_hex_chars => {
        r#"<?php
echo strlen(md5('abc'));
"#,
        ["32"]
    };

    sha1_string_produces_40_hex_chars => {
        r#"<?php
echo strlen(sha1('abc'));
"#,
        ["40"]
    };

    hash_sha256_known_digest => {
        r#"<?php
echo hash('sha256', 'abc');
"#,
        ["ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"]
    };

    hash_hmac_sha256_produces_64_hex_chars => {
        r#"<?php
echo strlen(hash_hmac('sha256', 'payload', 'secret-key'));
"#,
        ["64"]
    };

    hash_equals_true_for_matching_strings => {
        r#"<?php
echo hash_equals('abc', 'abc') ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    hash_equals_false_for_different_strings => {
        r#"<?php
echo hash_equals('abc', 'abd') ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    crc32_returns_non_zero_for_string => {
        r#"<?php
echo crc32('test') !== 0 ? 'crc' : 'zero';
"#,
        ["crc"]
    };

    bin2hex_hex2bin_roundtrip => {
        r#"<?php
echo hex2bin(bin2hex('ab'));
"#,
        ["ab"]
    };

    hash_algos_includes_sha256 => {
        r#"<?php
echo in_array('sha256', hash_algos(), true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    hash_file_from_memory_stream => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'data');
rewind($fp);
echo strlen(hash('md5', stream_get_contents($fp)));
"#,
        ["32"]
    };

    password_hash_and_verify_roundtrip => {
        r#"<?php
$h = password_hash('secret', PASSWORD_BCRYPT);
echo password_verify('secret', $h) ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    password_needs_rehash_false_for_fresh_hash => {
        r#"<?php
$h = password_hash('x', PASSWORD_BCRYPT);
echo password_needs_rehash($h, PASSWORD_BCRYPT) ? 'yes' : 'no';
"#,
        ["no"]
    };

    hash_hkdf_derives_key_bytes => {
        r#"<?php
echo strlen(hash_hkdf('sha256', 'input', 16, 'salt', 'info'));
"#,
        ["16"]
    };

    hash_pbkdf2_derives_hex_key => {
        r#"<?php
echo strlen(hash_pbkdf2('sha256', 'password', 'salt', 1000, 20));
"#,
        ["20"]
    };

    md5_file_equals_md5_of_contents => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'x');
rewind($fp);
$uri = stream_get_meta_data($fp)['uri'];
echo md5('x') === md5_file($uri) ? 'match' : 'diff';
"#,
        ["match"]
    };

    hash_hmac_algos_returns_array => {
        r#"<?php
echo is_array(hash_hmac_algos()) && in_array('sha256', hash_hmac_algos(), true) ? 'hmac_algos_ok' : 'err';
"#,
        ["hmac_algos_ok"]
    };

    hash_init_update_final_incremental => {
        r#"<?php
$ctx = hash_init('sha256');
hash_update($ctx, 'hello ');
hash_update($ctx, 'world');
echo hash_final($ctx);
"#,
        ["b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9"]
    };
}

