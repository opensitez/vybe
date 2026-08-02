<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_asymmetric_visibility_protected_set
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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

class BaseDocument {
    public protected(set) string $title = "Untitled";
}

class ArticleDocument extends BaseDocument {
    public function setTitle(string $title): void {
        $this->title = $title;
    }
}

$art = new ArticleDocument();
$art->setTitle("PHP 8.4 Released");
echo $art->title;

__vybe_check(ob_get_clean(), "PHP 8.4 Released");
