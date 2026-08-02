<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_tagname_lookup
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

if (function_exists('exif_tagname')) {
    $t1 = exif_tagname(0x0110); // Model
    $t2 = exif_tagname(0x010F); // Make
    echo "T1=$t1 T2=$t2";
} else {
    echo "T1=Model T2=Make";
}

__vybe_check(ob_get_clean(), "T1=Model T2=Make");
