//! `password_hash`, `password_verify`, `password_needs_rehash`, and `hash_*` password APIs.

crate::php_cases! {
    password_hash_bcrypt_prefix => {
        r#"<?php
$h = password_hash('secret', PASSWORD_BCRYPT);
echo str_starts_with($h, '$2y$') ? 'bcrypt' : 'other';
"#,
        ["bcrypt"]
    };

    password_verify_accepts_correct_password => {
        r#"<?php
$h = password_hash('right', PASSWORD_BCRYPT);
echo password_verify('right', $h) ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    password_verify_rejects_wrong_password => {
        r#"<?php
$h = password_hash('right', PASSWORD_BCRYPT);
echo password_verify('wrong', $h) ? 'ok' : 'fail';
"#,
        ["fail"]
    };

    password_needs_rehash_false_for_fresh_hash => {
        r#"<?php
$h = password_hash('x', PASSWORD_BCRYPT);
echo password_needs_rehash($h, PASSWORD_BCRYPT) ? 'yes' : 'no';
"#,
        ["no"]
    };

    password_get_info_identifies_bcrypt_algo => {
        r#"<?php
$h = password_hash('x', PASSWORD_BCRYPT);
$info = password_get_info($h);
echo ($info['algoName'] ?? '') === 'bcrypt' ? 'bcrypt' : 'other';
"#,
        ["bcrypt"]
    };

    password_hash_default_is_string => {
        r#"<?php
echo is_string(password_hash('pw', PASSWORD_DEFAULT)) ? 'str' : 'no';
"#,
        ["str"]
    };

    hash_equals_timing_safe_compare => {
        r#"<?php
echo hash_equals('abc', 'abc') ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    hash_hmac_sha256_length => {
        r#"<?php
echo strlen(hash_hmac('sha256', 'body', 'key'));
"#,
        ["64"]
    };

    sodium_crypto_generichash_when_extension_loaded => {
        r#"<?php
if (!function_exists('sodium_crypto_generichash')) { echo 'skip'; } else {
    echo strlen(sodium_crypto_generichash('data'));
}
"#,
        ["32"]
    };
}
