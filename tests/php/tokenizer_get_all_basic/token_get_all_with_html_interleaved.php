<?php
// vybe-test: php/tokenizer_get_all_basic/token_get_all_with_html_interleaved
// origin: languages/php/tests/php/test_tokenizer_get_all_basic.rs

$source = '<html><?php $a = 1; ?></html>';
$tokens = token_get_all($source);
$names = [];
foreach ($tokens as $t) {
    if (is_array($t)) {
        $names[] = token_name($t[0]);
    }
}
echo implode('|', $names);
