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

    // Character classes, including negation and ranges. Expected values
    // measured against real `php`. The pre-2026-08-01 emitter escaped `[` and
    // `]` as regex metacharacters, so every one of these matched the class
    // LITERALLY and returned false.
    fnmatch_character_classes => {
        r#"<?php
echo fnmatch("[!a]bc", "xbc") ? "1" : "0";
echo fnmatch("[!a]bc", "abc") ? "1" : "0";
echo fnmatch("[a-c]x", "bx") ? "1" : "0";
echo fnmatch("[a-c]x", "dx") ? "1" : "0";
"#,
        ["1010"]
    };

    // `.` is a regex metacharacter and must stay literal in a glob, while `?`
    // is the glob's own single-character wildcard.
    fnmatch_dot_is_literal_but_question_is_wildcard => {
        r#"<?php
echo fnmatch("a.b", "aXb") ? "1" : "0";
echo fnmatch("a.b", "a.b") ? "1" : "0";
echo fnmatch("a?b", "aXb") ? "1" : "0";
"#,
        ["011"]
    };
}
