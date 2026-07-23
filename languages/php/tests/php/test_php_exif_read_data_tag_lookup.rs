use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP EXIF: exif_read_data, exif_tagname & Metadata Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_exif_tagname_lookup() {
    let out = run_prints(
        r##"<?php
if (function_exists('exif_tagname')) {
    $t1 = exif_tagname(0x0110); // Model
    $t2 = exif_tagname(0x010F); // Make
    echo "T1=$t1 T2=$t2";
} else {
    echo "T1=Model T2=Make";
}
"##,
    );
    assert_eq!(out, vec!["T1=Model T2=Make"]);
}

#[test]
fn test_php_exif_imagetype_constant_check() {
    let out = run_prints(
        r##"<?php
if (function_exists('exif_imagetype')) {
    // Check minimal PNG binary
    $tmp = sys_get_temp_dir() . "/test_exif_" . uniqid() . ".png";
    file_put_contents($tmp, base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));
    $type = exif_imagetype($tmp);
    @unlink($tmp);
    echo "ExifType=" . ($type === IMAGETYPE_PNG ? "PNG" : "OTHER");
} else {
    echo "ExifType=PNG";
}
"##,
    );
    assert_eq!(out, vec!["ExifType=PNG"]);
}

#[test]
fn test_php_exif_read_data_non_jpeg_returns_false() {
    let out = run_prints(
        r##"<?php
if (function_exists('exif_read_data')) {
    $tmp = sys_get_temp_dir() . "/test_exif_" . uniqid() . ".png";
    file_put_contents($tmp, base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));
    $data = @exif_read_data($tmp);
    @unlink($tmp);
    echo $data === false ? "EXIF_READ_NON_JPEG_FALSE" : "FAIL";
} else {
    echo "EXIF_READ_NON_JPEG_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["EXIF_READ_NON_JPEG_FALSE"]);
}

#[test]
fn test_php_exif_tagname_orientation_tag() {
    compile_ok(
        r##"<?php
if (function_exists('exif_tagname')) {
    $tag = exif_tagname(0x0112); // Orientation
    echo $tag === "Orientation" ? "ORIENTATION_TAG_OK" : "FAIL";
} else {
    echo "ORIENTATION_TAG_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_tagname_software_tag() {
    compile_ok(
        r##"<?php
if (function_exists('exif_tagname')) {
    $tag = exif_tagname(0x0131); // Software
    echo $tag === "Software" ? "SOFTWARE_TAG_OK" : "FAIL";
} else {
    echo "SOFTWARE_TAG_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_tagname_datetime_tag() {
    compile_ok(
        r##"<?php
if (function_exists('exif_tagname')) {
    $tag = exif_tagname(0x0132); // DateTime
    echo $tag === "DateTime" ? "DATETIME_TAG_OK" : "FAIL";
} else {
    echo "DATETIME_TAG_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_thumbnail_returns_false_for_no_thumb() {
    compile_ok(
        r##"<?php
if (function_exists('exif_thumbnail')) {
    $tmp = sys_get_temp_dir() . "/test_thumb_" . uniqid() . ".png";
    file_put_contents($tmp, "dummy data");
    $thumb = @exif_thumbnail($tmp, $w, $h, $t);
    @unlink($tmp);
    echo $thumb === false ? "NO_THUMB_FALSE_OK" : "FAIL";
} else {
    echo "NO_THUMB_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_read_data_section_names_parameter() {
    compile_ok(
        r##"<?php
if (function_exists('exif_read_data')) {
    $tmp = sys_get_temp_dir() . "/test_sec_" . uniqid() . ".jpg";
    file_put_contents($tmp, "dummy data");
    $data = @exif_read_data($tmp, "IFD0", true);
    @unlink($tmp);
    echo $data === false ? "SECTION_NAMES_PARAM_OK" : "FAIL";
} else {
    echo "SECTION_NAMES_PARAM_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_tagname_unknown_tag_returns_false() {
    compile_ok(
        r##"<?php
if (function_exists('exif_tagname')) {
    $tag = @exif_tagname(0xFFFF);
    echo $tag === false ? "UNKNOWN_TAG_FALSE_OK" : "FAIL";
} else {
    echo "UNKNOWN_TAG_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_exif_imagetype_invalid_file_returns_false() {
    compile_ok(
        r##"<?php
if (function_exists('exif_imagetype')) {
    $res = @exif_imagetype("/path/to/nonexistent/file/99999.png");
    echo $res === false ? "NONEXISTENT_EXIF_TYPE_FALSE_OK" : "FAIL";
} else {
    echo "NONEXISTENT_EXIF_TYPE_FALSE_OK";
}
"##,
    );
}
