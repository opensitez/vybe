use super::helpers::run_prints;

#[test]
fn test_strip_tags_html_comment_removal() {
    assert_eq!(
        run_prints(
            r#"<?php
$html = "<!-- secret comment -->Hello <b>World</b>";
echo strip_tags($html), "\n";
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn test_strip_tags_unclosed_tag_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
$html = "Text <p>Paragraph <a href=";
echo strip_tags($html), "\n";
"#
        ),
        vec!["Text Paragraph "]
    );
}
