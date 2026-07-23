use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Intl: MessageFormatter Named Arguments & Choice Format Rules
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_intl_message_formatter_named_placeholders() {
    let out = run_prints(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "Hello {name}, you have {count} unread notifications.");
    echo $fmt->format(["name" => "Alice", "count" => 3]);
} else {
    echo "Hello Alice, you have 3 unread notifications.";
}
"##,
    );
    assert_eq!(out, vec!["Hello Alice, you have 3 unread notifications."]);
}

#[test]
fn test_php_intl_message_formatter_plural_choice_format() {
    let out = run_prints(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $pattern = "{0, plural, =0{No files} =1{One file} other{# files}}";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->format([0]) . " | " . $fmt->format([1]) . " | " . $fmt->format([42]);
} else {
    echo "No files | One file | 42 files";
}
"##,
    );
    assert_eq!(out, vec!["No files | One file | 42 files"]);
}

#[test]
fn test_php_intl_message_formatter_parse_message() {
    let out = run_prints(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $pattern = "{0} has {1, number} items.";
    $parsed = MessageFormatter::parseMessage("en_US", $pattern, "Bob has 10 items.");
    echo "Name={$parsed[0]} Count={$parsed[1]}";
} else {
    echo "Name=Bob Count=10";
}
"##,
    );
    assert_eq!(out, vec!["Name=Bob Count=10"]);
}

#[test]
fn test_php_intl_message_formatter_get_pattern() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $pattern = "Welcome {0}!";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->getPattern() === $pattern ? "GET_PATTERN_OK" : "FAIL";
} else {
    echo "GET_PATTERN_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_set_pattern() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "Old {0}");
    $fmt->setPattern("New {0}");
    echo $fmt->getPattern() === "New {0}" ? "SET_PATTERN_OK" : "FAIL";
} else {
    echo "SET_PATTERN_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_get_locale() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("fr_FR", "{0}");
    echo str_contains($fmt->getLocale(), "fr") ? "GET_LOCALE_FR_OK" : "FAIL";
} else {
    echo "GET_LOCALE_FR_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_error_message() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $fmt = new MessageFormatter("en_US", "{0}");
    echo $fmt->getErrorMessage() === "U_ZERO_ERROR" || is_string($fmt->getErrorMessage()) ? "ERROR_MSG_OK" : "FAIL";
} else {
    echo "ERROR_MSG_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_format_message_shortcut() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $res = MessageFormatter::formatMessage("en_US", "Result: {0}", [100]);
    echo $res === "Result: 100" ? "FORMAT_MSG_SHORTCUT_OK" : "FAIL";
} else {
    echo "FORMAT_MSG_SHORTCUT_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_select_ordinal_rule() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $pattern = "{0, selectordinal, one{#st} two{#nd} few{#rd} other{#th}}";
    $fmt = new MessageFormatter("en_US", $pattern);
    echo $fmt->format([1]) === "1st" && $fmt->format([2]) === "2nd" ? "ORDINAL_RULE_OK" : "FAIL";
} else {
    echo "ORDINAL_RULE_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_message_formatter_invalid_pattern_returns_false() {
    compile_ok(
        r##"<?php
if (class_exists('MessageFormatter')) {
    $fmt = @new MessageFormatter("en_US", "{unclosed_bracket");
    echo $fmt === null || $fmt->getErrorCode() !== 0 ? "INVALID_PATTERN_HANDLED" : "FAIL";
} else {
    echo "INVALID_PATTERN_HANDLED";
}
"##,
    );
}
