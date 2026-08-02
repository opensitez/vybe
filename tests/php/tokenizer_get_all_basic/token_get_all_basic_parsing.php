<?php
// vybe-test: php/tokenizer_get_all_basic/token_get_all_basic_parsing
// origin: languages/php/tests/php/test_tokenizer_get_all_basic.rs

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
