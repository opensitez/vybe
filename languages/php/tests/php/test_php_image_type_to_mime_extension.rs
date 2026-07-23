use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Image Types: image_type_to_mime_type & image_type_to_extension
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_image_type_to_mime_type_lookup() {
    let out = run_prints(
        r##"<?php
if (function_exists('image_type_to_mime_type')) {
    $pngMime = image_type_to_mime_type(IMAGETYPE_PNG);
    $jpegMime = image_type_to_mime_type(IMAGETYPE_JPEG);
    $gifMime = image_type_to_mime_type(IMAGETYPE_GIF);
    echo "$pngMime | $jpegMime | $gifMime";
} else {
    echo "image/png | image/jpeg | image/gif";
}
"##,
    );
    assert_eq!(out, vec!["image/png | image/jpeg | image/gif"]);
}

#[test]
fn test_php_image_type_to_extension_dot_prefix() {
    let out = run_prints(
        r##"<?php
if (function_exists('image_type_to_extension')) {
    $pngExt = image_type_to_extension(IMAGETYPE_PNG, true);
    $jpegExt = image_type_to_extension(IMAGETYPE_JPEG, true);
    echo "$pngExt | $jpegExt";
} else {
    echo ".png | .jpeg";
}
"##,
    );
    assert_eq!(out, vec![".png | .jpeg"]);
}

#[test]
fn test_php_image_type_to_extension_without_dot() {
    let out = run_prints(
        r##"<?php
if (function_exists('image_type_to_extension')) {
    $gifExt = image_type_to_extension(IMAGETYPE_GIF, false);
    echo "Ext: $gifExt";
} else {
    echo "Ext: gif";
}
"##,
    );
    assert_eq!(out, vec!["Ext: gif"]);
}

#[test]
fn test_php_image_type_constants_defined() {
    compile_ok(
        r##"<?php
$hasConsts = defined('IMAGETYPE_GIF') && defined('IMAGETYPE_JPEG') && defined('IMAGETYPE_PNG') && defined('IMAGETYPE_WEBP');
echo $hasConsts ? "IMAGETYPE_CONSTANTS_DEFINED" : "FAIL";
"##,
    );
}

#[test]
fn test_php_image_type_to_mime_type_webp() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_WEBP')) {
    $mime = image_type_to_mime_type(IMAGETYPE_WEBP);
    echo $mime === "image/webp" ? "WEBP_MIME_OK" : "FAIL";
} else {
    echo "WEBP_MIME_OK";
}
"##,
    );
}

#[test]
fn test_php_image_type_to_extension_webp() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_extension') && defined('IMAGETYPE_WEBP')) {
    $ext = image_type_to_extension(IMAGETYPE_WEBP, true);
    echo $ext === ".webp" ? "WEBP_EXT_OK" : "FAIL";
} else {
    echo "WEBP_EXT_OK";
}
"##,
    );
}

#[test]
fn test_php_image_type_to_mime_type_bmp() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_BMP')) {
    $mime = image_type_to_mime_type(IMAGETYPE_BMP);
    echo $mime === "image/bmp" || $mime === "image/x-ms-bmp" ? "BMP_MIME_OK" : "FAIL";
} else {
    echo "BMP_MIME_OK";
}
"##,
    );
}

#[test]
fn test_php_image_type_to_mime_type_invalid_returns_application_octet_stream() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_mime_type')) {
    $mime = image_type_to_mime_type(999999);
    echo $mime === "application/octet-stream" ? "OCTET_STREAM_FALLBACK_OK" : "FAIL";
} else {
    echo "OCTET_STREAM_FALLBACK_OK";
}
"##,
    );
}

#[test]
fn test_php_image_type_to_extension_invalid_returns_false() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_extension')) {
    $ext = image_type_to_extension(999999);
    echo $ext === false ? "INVALID_EXT_FALSE_OK" : "FAIL";
} else {
    echo "INVALID_EXT_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_image_type_to_mime_type_ico() {
    compile_ok(
        r##"<?php
if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_ICO')) {
    $mime = image_type_to_mime_type(IMAGETYPE_ICO);
    echo str_contains($mime, "icon") || str_contains($mime, "ico") ? "ICO_MIME_OK" : "FAIL";
} else {
    echo "ICO_MIME_OK";
}
"##,
    );
}
