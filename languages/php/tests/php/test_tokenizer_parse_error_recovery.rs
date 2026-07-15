use super::helpers::run_prints;

crate::php_cases! {
    token_get_all_parse_error => {
        r#"<?php
// token_get_all should silently tokenize even if there is a parse error
$source = '<?php class { public }';
$tokens = token_get_all($source);
$count = 0;
foreach ($tokens as $t) {
    if (is_array($t) && (token_name($t[0]) === 'T_CLASS' || token_name($t[0]) === 'T_PUBLIC')) {
        $count++;
    }
}
echo $count;
"#,
        ["2"]
    };

    token_get_all_token_parse_flag => {
        r#"<?php
// In PHP 8+, TOKEN_PARSE flag can throw ParseError on invalid syntax
$source = '<?php class { public }';
try {
    token_get_all($source, TOKEN_PARSE);
    echo "success";
} catch (\ParseError $e) {
    echo "error";
}
"#,
        ["error"]
    };
}
