
crate::php_cases! {
    substr_compare_basic => {
        r#"<?php
echo substr_compare("abcde", "bc", 1, 2);
"#,
        ["0"]
    };

    substr_compare_case_insensitivity => {
        r#"<?php
echo substr_compare("abcde", "BC", 1, 2, true);
"#,
        ["0"]
    };

    substr_compare_negative_offset => {
        r#"<?php
echo substr_compare("abcde", "de", -2, 2);
"#,
        ["0"]
    };
}
