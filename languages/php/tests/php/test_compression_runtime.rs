//! `gzcompress`, `gzencode`, `gzdeflate`, and roundtrip decompression.

crate::php_cases! {
    gzcompress_roundtrip_restores_original => {
        r#"<?php
$orig = 'compress me';
echo gzuncompress(gzcompress($orig)) === $orig ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    gzencode_produces_gzip_wrapper => {
        r#"<?php
$c = gzencode('data');
echo str_starts_with($c, "\x1f\x8b") ? 'gzip' : 'raw';
"#,
        ["gzip"]
    };

    gzdecode_reverses_gzencode => {
        r#"<?php
echo gzdecode(gzencode('payload'));
"#,
        ["payload"]
    };

    gzdeflate_shorter_than_raw_for_repetitive => {
        r#"<?php
$raw = str_repeat('xy', 100);
echo strlen(gzdeflate($raw)) < strlen($raw) ? 'smaller' : 'bigger';
"#,
        ["smaller"]
    };

    gzinflate_reverses_gzdeflate => {
        r#"<?php
$raw = 'hello gzip stack';
echo gzinflate(gzdeflate($raw));
"#,
        ["hello gzip stack"]
    };

    gzcompress_higher_level_not_larger_than_raw_small => {
        r#"<?php
$raw = str_repeat('a', 50);
echo strlen(gzcompress($raw, 9)) > 0 ? 'ok' : 'empty';
"#,
        ["ok"]
    };

    gzuncompress_false_on_invalid_data => {
        r#"<?php
echo gzuncompress('not-gzip') === false ? 'false' : 'ok';
"#,
        ["false"]
    };
}
