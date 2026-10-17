<?php
function ordinary_records() { yield "a" => 3; yield 7; }
function collected_values() { return iterator_to_array(ordinary_records(), false); }
function collected_keys() { return iterator_to_array(ordinary_records(), true); }
function nullable_serialized() { return serialize(null); }
function scalar_serialized() { return serialize(3); }

function native_in_array_controls() { return [in_array(3,[3,4],true),in_array(5,[3,4],true),in_array("x",["x","y"],true),in_array("z",["x","y"],true),in_array("3",[3],false),in_array("z",[3],false)]; }

function native_pathinfo_controls() { return [pathinfo("/a/b.txt",1),pathinfo("/a/b.txt",2),pathinfo("/a/b.txt",4),pathinfo("/a/b.txt",8),pathinfo("file",4)]; }

function native_ini_controls() { return [ini_get("precision"),ini_get("unknown-probe"),ini_set("precision","!"),ini_get("precision"),ini_restore("precision"),ini_get("precision")]; }

function native_header_controls() { header("HTTP/1.1 201 Created"); header("HTTP/1.1 204"); header("X-Probe: yes",true); header("X-Probe: again",false); return null; }

function native_echo_controls() { ob_start(); echo "a",3; $n=print "b"; $s=ob_get_clean(); return [$s,$n]; }

function native_random_controls() { return [mt_rand(7,7),rand(1,3),lcg_value()]; }

function native_file_size_control($path) { return filesize($path); }

function native_callback($a,$b) { return $a+$b; }
function native_callback_controls() { return [call_user_func("native_callback",3,4),call_user_func_array("native_callback",[5,6])]; }
