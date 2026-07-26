
crate::php_cases! {
    token_get_all_basic_parsing => {
        r#"<?php
$source = '<?php echo "hello"; ?>';
$tokens = token_get_all($source);
$output = [];
foreach ($tokens as $token) {
    if (is_array($token)) {
        $output[] = token_name($token[0]);
    } else {
        $output[] = $token;
    }
}
echo implode(',', $output);
"#,
        ["T_OPEN_TAG,T_ECHO,T_WHITESPACE,T_CONSTANT_ENCAPSED_STRING,;,T_WHITESPACE,T_CLOSE_TAG"]
    };

    token_get_all_with_html_interleaved => {
        r#"<?php
$source = '<html><?php $a = 1; ?></html>';
$tokens = token_get_all($source);
$names = [];
foreach ($tokens as $t) {
    if (is_array($t)) {
        $names[] = token_name($t[0]);
    }
}
echo implode('|', $names);
"#,
        ["T_INLINE_HTML|T_OPEN_TAG|T_VARIABLE|T_WHITESPACE|T_LNUMBER|T_WHITESPACE|T_CLOSE_TAG|T_INLINE_HTML"]
    };
}
