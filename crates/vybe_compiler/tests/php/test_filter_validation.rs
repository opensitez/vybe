//! `filter_var` validation and sanitization — distinct filter IDs and flags.

crate::php_cases! {
    filter_validate_email_accepts_user_at_host => {
        r#"<?php
echo filter_var('user@example.com', FILTER_VALIDATE_EMAIL) !== false ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    filter_validate_email_rejects_missing_domain => {
        r#"<?php
echo filter_var('user@', FILTER_VALIDATE_EMAIL) === false ? 'bad' : 'ok';
"#,
        ["bad"]
    };

    filter_validate_url_https => {
        r#"<?php
echo filter_var('https://example.com/path', FILTER_VALIDATE_URL) !== false ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    filter_validate_url_rejects_spaces => {
        r#"<?php
echo filter_var('not a url', FILTER_VALIDATE_URL) === false ? 'bad' : 'ok';
"#,
        ["bad"]
    };

    filter_validate_ip_v4_ok => {
        r#"<?php
echo filter_var('192.168.0.1', FILTER_VALIDATE_IP) !== false ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    filter_validate_ip_rejects_256_octet => {
        r#"<?php
echo filter_var('256.0.0.1', FILTER_VALIDATE_IP) === false ? 'bad' : 'ok';
"#,
        ["bad"]
    };

    filter_validate_ip_v6_flag => {
        r#"<?php
echo filter_var('::1', FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) !== false ? 'v6' : 'no';
"#,
        ["v6"]
    };

    filter_validate_int_accepts_digits => {
        r#"<?php
echo filter_var('42', FILTER_VALIDATE_INT);
"#,
        ["42"]
    };

    filter_validate_int_rejects_float_string => {
        r#"<?php
echo filter_var('3.14', FILTER_VALIDATE_INT) === false ? 'bad' : 'ok';
"#,
        ["bad"]
    };

    filter_validate_float_accepts_decimal => {
        r#"<?php
echo filter_var('3.14', FILTER_VALIDATE_FLOAT);
"#,
        ["3.14"]
    };

    filter_validate_boolean_true_strings => {
        r#"<?php
echo filter_var('true', FILTER_VALIDATE_BOOLEAN) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    filter_validate_boolean_false_for_zero => {
        r#"<?php
echo filter_var('0', FILTER_VALIDATE_BOOLEAN) ? 'yes' : 'no';
"#,
        ["no"]
    };

    filter_sanitize_email_strips_invalid_chars => {
        // PHP keeps '!' — it is inside the FILTER_SANITIZE_EMAIL allow-set;
        // only the space is stripped. Verified against PHP 8.4 CLI.
        r#"<?php
echo filter_var('bad email!', FILTER_SANITIZE_EMAIL);
"#,
        ["bademail!"]
    };

    filter_sanitize_url_removes_invalid => {
        r#"<?php
echo str_starts_with(filter_var(' https://ex.com ', FILTER_SANITIZE_URL), 'https') ? 'url' : 'no';
"#,
        ["url"]
    };

    filter_sanitize_number_int_strips_letters => {
        r#"<?php
echo filter_var('a1b2c3', FILTER_SANITIZE_NUMBER_INT);
"#,
        ["123"]
    };

    filter_sanitize_string_strips_tags => {
        r#"<?php
echo strip_tags('<b>x</b>');
"#,
        ["x"]
    };

    filter_var_array_validates_multiple => {
        r#"<?php
$in = ['a' => '1', 'b' => 'x'];
$out = filter_var_array($in, ['a' => FILTER_VALIDATE_INT, 'b' => FILTER_VALIDATE_INT]);
echo ($out['a'] === 1 ? '1' : '0') . ($out['b'] === false ? 'f' : 't');
"#,
        ["1f"]
    };

    filter_has_var_checks_get_superglobal => {
        // filter_has_var(INPUT_GET, …) inspects the real request input, not a
        // runtime-mutated $_GET; under the CLI there is none, so PHP returns
        // false → 'no'. Verified against PHP 8.4 CLI.
        r#"<?php
$_GET['probe'] = '1';
echo filter_has_var(INPUT_GET, 'probe') ? 'has' : 'no';
"#,
        ["no"]
    };

    filter_list_returns_available_filters => {
        // filter_list() returns filter *names* (strings); comparing the int id
        // FILTER_VALIDATE_EMAIL with strict in_array is always false → 'missing'.
        // Verified against PHP 8.4 CLI.
        r#"<?php
echo in_array(FILTER_VALIDATE_EMAIL, filter_list(), true) ? 'listed' : 'missing';
"#,
        ["missing"]
    };

    filter_id_by_name_validate_email => {
        r#"<?php
echo filter_id('validate_email') === FILTER_VALIDATE_EMAIL ? 'match' : 'diff';
"#,
        ["match"]
    };
}
