<?php
// vybe-test: php/oop_advanced/named_args_skip_defaults
// origin: languages/php/tests/php/test_oop_advanced.rs

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

function createTag(string $tag, string $content, string $class = "", string $id = ""): string {
    $attrs = "";
    if ($class) $attrs .= " class=\"$class\"";
    if ($id) $attrs .= " id=\"$id\"";
    return "<$tag$attrs>$content</$tag>";
}
echo createTag("div", "hello", id: "main"), "\n";
echo createTag(tag: "span", content: "world", class: "bold"), "\n";

__vybe_check(ob_get_clean(), "<div id=\"main\">hello</div>\n<span class=\"bold\">world</span>");
