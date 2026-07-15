use super::helpers::run_prints;

crate::php_cases! {
    wordwrap_basic => {
        r#"<?php
$str = "The quick brown fox jumped over the lazy dog.";
echo wordwrap($str, 20, "<br>\n");
"#,
        ["The quick brown fox<br>\njumped over the lazy<br>\ndog."]
    };

    wordwrap_cut_long_words => {
        r#"<?php
$str = "A very long woooooooooooord.";
echo wordwrap($str, 8, "\n", true);
"#,
        ["A very\nlong\nwooooooo\nooooord."]
    };
}
