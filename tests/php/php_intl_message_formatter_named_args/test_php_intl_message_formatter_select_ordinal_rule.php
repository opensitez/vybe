<?php
// vybe-test: php/php_intl_message_formatter_named_args/test_php_intl_message_formatter_select_ordinal_rule
// origin: languages/php/tests/php/test_php_intl_message_formatter_named_args.rs
// vybe-test-mode: compile

if (class_exists('MessageFormatter')) {
    $pattern = "{0, selectordinal, one{#st} two{#nd} few{#rd} other{#th}}";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->format([1]) === "1st" && $fmt->format([2]) === "2nd" ? "ORDINAL_RULE_OK" : "FAIL";
} else {
    echo "ORDINAL_RULE_OK";
}
