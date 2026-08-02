<?php
// vybe-test: php/type_juggling/juggling_in_switch
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$val = "0";
switch ($val) {
    case false: echo "false"; break;
    case null:  echo "null";  break;
    case 0:     echo "zero";  break;
    case "0":   echo "string zero"; break;
    default:    echo "default";
}
