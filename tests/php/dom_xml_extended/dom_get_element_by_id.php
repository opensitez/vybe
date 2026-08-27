<?php
// vybe-test: php/dom_xml_extended/dom_get_element_by_id
// origin: languages/php/tests/php/test_dom_xml_extended.rs

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

echo "dom_get_element_by_id_ok";

__vybe_check(ob_get_clean(), "dom_get_element_by_id_ok");
