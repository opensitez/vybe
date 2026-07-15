use super::helpers::run_prints;

crate::php_cases! {
    strip_tags_basic => {
        r#"<?php
$str = "<p>Test <b>paragraph</b>.</p><!-- Comment --> <a href='#'>Link</a>";
echo strip_tags($str);
"#,
        ["Test paragraph. Link"]
    };

    strip_tags_allowed_string => {
        r#"<?php
$str = "<p>Test <b>paragraph</b>.</p>";
echo strip_tags($str, "<b>");
"#,
        ["Test <b>paragraph</b>."]
    };

    strip_tags_allowed_array => {
        r#"<?php
$str = "<p>Test <b>paragraph</b>.</p> <i>italic</i>";
echo strip_tags($str, ['b', 'i']);
"#,
        ["Test <b>paragraph</b>. <i>italic</i>"]
    };
}
