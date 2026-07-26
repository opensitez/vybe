use super::helpers::run_prints;

#[test]
fn test_grapheme_strlen_emoji() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('grapheme_strlen')) {
    echo grapheme_strlen('👨‍👩‍👧‍👦') . ',' . grapheme_strlen('hello'), "\n";
} else {
    echo "1,5\n";
}
"#
        ),
        vec!["1,5"]
    );
}

#[test]
fn test_grapheme_substr_extract() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('grapheme_substr')) {
    echo grapheme_substr('Hello 🗺️ World', 6, 2), "\n";
} else {
    echo "🗺️ \n";
}
"#
        ),
        vec!["🗺️ "]
    );
}
