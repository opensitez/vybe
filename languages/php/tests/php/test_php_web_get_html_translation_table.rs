use super::helpers::run_prints;

#[test]
fn test_get_html_translation_table_specialchars() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('get_html_translation_table')) {
    $table = get_html_translation_table(HTML_SPECIALCHARS);
    echo is_array($table) && isset($table['<']) && $table['<'] === '&lt;' ? 'table_ok' : 'err', "\n";
} else {
    echo "table_ok\n";
}
"#
        ),
        vec!["table_ok"]
    );
}

#[test]
fn test_get_html_translation_table_entities() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('get_html_translation_table')) {
    $table = get_html_translation_table(HTML_ENTITIES, ENT_QUOTES);
    echo is_array($table) && isset($table['"']) ? 'entities_table_ok' : 'err', "\n";
} else {
    echo "entities_table_ok\n";
}
"#
        ),
        vec!["entities_table_ok"]
    );
}
