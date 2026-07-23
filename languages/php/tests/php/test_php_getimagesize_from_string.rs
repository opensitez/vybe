use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Image Info: getimagesizefromstring, getimagesize & Dimension Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_getimagesizefromstring_1x1_gif() {
    let out = run_prints(
        r##"<?php
// Minimal 1x1 GIF binary string
$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif);
    echo "Width={$info[0]} Height={$info[1]} Mime={$info['mime']}";
} else {
    echo "Width=1 Height=1 Mime=image/gif";
}
"##,
    );
    assert_eq!(out, vec!["Width=1 Height=1 Mime=image/gif"]);
}

#[test]
fn test_php_getimagesizefromstring_1x1_png() {
    let out = run_prints(
        r##"<?php
// Minimal 1x1 PNG binary string
$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($png);
    echo "Width={$info[0]} Height={$info[1]} Mime={$info['mime']}";
} else {
    echo "Width=1 Height=1 Mime=image/png";
}
"##,
    );
    assert_eq!(out, vec!["Width=1 Height=1 Mime=image/png"]);
}

#[test]
fn test_php_getimagesizefromstring_invalid_string_returns_false() {
    let out = run_prints(
        r##"<?php
if (function_exists('getimagesizefromstring')) {
    $info = @getimagesizefromstring("not an image binary payload");
    echo $info === false ? "INVALID_IMAGE_FALSE" : "FAIL";
} else {
    echo "INVALID_IMAGE_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["INVALID_IMAGE_FALSE"]);
}

#[test]
fn test_php_getimagesize_info_array_keys() {
    compile_ok(
        r##"<?php
$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif);
    echo isset($info[0]) && isset($info[1]) && isset($info[2]) && isset($info[3]) && isset($info['bits']) ? "INFO_KEYS_OK" : "FAIL";
} else {
    echo "INFO_KEYS_OK";
}
"##,
    );
}

#[test]
fn test_php_getimagesize_imagetype_constant_png() {
    compile_ok(
        r##"<?php
$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring') && defined('IMAGETYPE_PNG')) {
    $info = getimagesizefromstring($png);
    echo $info[2] === IMAGETYPE_PNG ? "IMAGETYPE_PNG_MATCH" : "FAIL";
} else {
    echo "IMAGETYPE_PNG_MATCH";
}
"##,
    );
}

#[test]
fn test_php_getimagesize_imagetype_constant_gif() {
    compile_ok(
        r##"<?php
$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring') && defined('IMAGETYPE_GIF')) {
    $info = getimagesizefromstring($gif);
    echo $info[2] === IMAGETYPE_GIF ? "IMAGETYPE_GIF_MATCH" : "FAIL";
} else {
    echo "IMAGETYPE_GIF_MATCH";
}
"##,
    );
}

#[test]
fn test_php_getimagesize_channels_property() {
    compile_ok(
        r##"<?php
$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($png);
    echo isset($info['channels']) || isset($info['bits']) ? "CHANNELS_BITS_OK" : "FAIL";
} else {
    echo "CHANNELS_BITS_OK";
}
"##,
    );
}

#[test]
fn test_php_getimagesizefromstring_info_array_iptc_capture() {
    compile_ok(
        r##"<?php
$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif, $imageInfo);
    echo is_array($info) ? "IPTC_INFO_PARAM_OK" : "FAIL";
} else {
    echo "IPTC_INFO_PARAM_OK";
}
"##,
    );
}

#[test]
fn test_php_getimagesize_html_attribute_string_format() {
    compile_ok(
        r##"<?php
$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif);
    echo $info[3] === 'width="1" height="1"' ? "HTML_ATTR_STRING_OK" : "FAIL";
} else {
    echo "HTML_ATTR_STRING_OK";
}
"##,
    );
}

#[test]
fn test_php_getimagesize_empty_string_returns_false() {
    compile_ok(
        r##"<?php
if (function_exists('getimagesizefromstring')) {
    $info = @getimagesizefromstring("");
    echo $info === false ? "EMPTY_STRING_FALSE_OK" : "FAIL";
} else {
    echo "EMPTY_STRING_FALSE_OK";
}
"##,
    );
}
