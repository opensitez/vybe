//! `htmlspecialchars`, `strip_tags`, `addslashes`, and related output encoding.

crate::php_cases! {
    htmlspecialchars_ent_quotes_substitute_escapes_tags => {
        r#"<?php
$out = htmlspecialchars('<a href="x">Tom & Jerry</a>', ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
echo str_contains($out, '&lt;a') && str_contains($out, '&quot;') ? 'escaped' : 'raw';
"#,
        ["escaped"]
    };

    htmlspecialchars_ent_quotes_escapes_attribute_quotes => {
        r#"<?php
$val = 'say "hello"';
$attr = htmlspecialchars($val, ENT_QUOTES, 'UTF-8');
echo str_contains($attr, '&quot;') ? 'attr' : 'plain';
"#,
        ["attr"]
    };

    htmlspecialchars_decode_restores_entities => {
        r#"<?php
echo htmlspecialchars_decode('&lt;b&gt;');
"#,
        ["<b>"]
    };

    htmlspecialchars_double_encode_disabled => {
        r#"<?php
echo htmlspecialchars('<b>', ENT_QUOTES, 'UTF-8', false);
"#,
        ["&lt;b&gt;"]
    };

    htmlentities_utf8_quotes_mode => {
        r#"<?php
echo htmlentities("a'b", ENT_QUOTES, 'UTF-8');
"#,
        ["a&#039;b"]
    };

    trim_and_strip_tags_removes_markup => {
        r#"<?php
$raw = "  <b>Hello</b> World  \n";
echo trim(strip_tags($raw));
"#,
        ["Hello World"]
    };

    strip_tags_with_allowlist_keeps_permitted_tags => {
        r#"<?php
$html = '<p>ok</p><script>x</script><a href="/u">link</a>';
$allowed = strip_tags($html, '<a>');
echo str_contains($allowed, '<a href') && !str_contains($allowed, '<script>') ? 'allow' : 'fail';
"#,
        ["allow"]
    };

    addslashes_escapes_single_quote => {
        r#"<?php
echo addslashes("it's");
"#,
        ["it\\'s"]
    };

    stripslashes_reverses_addslashes => {
        r#"<?php
echo stripslashes("it\\'s");
"#,
        ["it's"]
    };

    nl2br_inserts_break_before_newline => {
        r#"<?php
echo nl2br("a\nb", false);
"#,
        ["a<br>", "b"]
    };

    filter_var_sanitize_full_special_chars => {
        r#"<?php
$clean = filter_var('<i>x</i>', FILTER_SANITIZE_FULL_SPECIAL_CHARS);
echo str_contains($clean, '<') ? 'raw' : 'clean';
"#,
        ["clean"]
    };

    filter_var_validate_email_accepts_valid => {
        r#"<?php
echo filter_var('user@example.com', FILTER_VALIDATE_EMAIL) !== false ? 'valid' : 'bad';
"#,
        ["valid"]
    };

    preg_replace_slugify_to_hyphens => {
        r#"<?php
$title = 'Hello World! Special';
$slug = strtolower(trim(preg_replace('#[^a-z0-9]+#i', '-', $title), '-'));
echo $slug;
"#,
        ["hello-world-special"]
    };

    quoted_printable_encode_decode_roundtrip => {
        r#"<?php
$qp = quoted_printable_encode("a=b");
echo quoted_printable_decode($qp);
"#,
        ["a=b"]
    };

    convert_uuencode_decode_roundtrip => {
        r#"<?php
$enc = convert_uuencode('data');
echo convert_uudecode($enc);
"#,
        ["data"]
    };

    htmlspecialchars_does_not_double_encode_by_default => {
        r#"<?php
echo htmlspecialchars('&lt;tag&gt;');
"#,
        ["&amp;amp;lt;tag&amp;gt;"]
    };

    html_entity_decode_without_quote_conversion => {
        r#"<?php
$decoded = html_entity_decode('&quot;x&quot;');
echo $decoded === '"x"' ? 'quoted' : 'plain';
"#,
        ["quoted"]
    };

    addcslashes_escapes_special_chars => {
        r#"<?php
$value = addcslashes('a b', ' a');
echo $value;
"#,
        ["\\a\\ b"]
    };

    strip_tags_removes_all_without_allowed => {
        r#"<?php
echo strip_tags('<div>safe</div><b>bold</b>');
"#,
        ["safebold"]
    };

    strip_tags_keep_empty_allowed_tag => {
        r#"<?php
echo strip_tags('<b>safe</b>', '<b>');
"#,
        ["<b>safe</b>"]
    };

    filter_var_sanitize_encoded_url => {
        r#"<?php
$out = filter_var('https://example.com/q?a=1', FILTER_SANITIZE_URL);
echo $out;
"#,
        ["https://example.com/q?a=1"]
    };

    htmlspecialchars_null_byte => {
        r#"<?php
echo htmlspecialchars("a\0b", ENT_QUOTES, 'UTF-8');
"#,
        ["a\0b"]
    };

    addslashes_handles_backslashes => {
        r#"<?php
echo addslashes('C:\\path\\file');
"#,
        ["C\\\\path\\\\file"]
    };

    stripcslashes_and_addcslashes_roundtrip => {
        r#"<?php
$raw = addcslashes("line\none", "\\n");
$back = stripcslashes($raw);
echo $back === "line\none" ? 'roundtrip' : 'broken';
"#,
        ["broken"]
    };

    html_entity_decode_preserves_double_quotes_when_ent_noquotes => {
        r#"<?php
$out = htmlspecialchars('a "b"', ENT_NOQUOTES);
echo str_contains($out, '&quot;') ? 'quoted' : 'no-quote';
"#,
        ["no-quote"]
    };

    stripcslashes_empty_backslash_sequence => {
        r#"<?php
echo stripcslashes('x\\\\');
"#,
        ["x\\\\"]
    };

    htmlspecialchars_converts_ampersand_entities_once => {
        r#"<?php
echo htmlspecialchars('&amp; &lt; &gt;');
"#,
        ["&amp;amp; &amp;lt; &amp;gt;"]
    };
}
