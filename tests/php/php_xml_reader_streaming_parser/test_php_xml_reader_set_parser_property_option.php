<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_set_parser_property_option
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_php_xml_reader_set_parser_property_option_ok";

__vybe_check(ob_get_clean(), "test_php_xml_reader_set_parser_property_option_ok");
