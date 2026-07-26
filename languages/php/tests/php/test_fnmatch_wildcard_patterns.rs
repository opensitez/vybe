
crate::php_cases! {
    fnmatch_basic => {
        r#"<?php
echo fnmatch("*gr[ae]y", "color_is_grey") ? "match|" : "no|";
echo fnmatch("*.php", "index.html") ? "match" : "no";
"#,
        ["match|no"]
    };

    fnmatch_flags => {
        r#"<?php
echo fnmatch("*gr[ae]y", "color_is_Grey", FNM_CASEFOLD) ? "match" : "no";
"#,
        ["match"]
    };
}
