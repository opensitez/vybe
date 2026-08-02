<?php
// vybe-test: php/php_web_htmlentities_charset_quotes/test_htmlentities_ent_quotes_mode
// origin: languages/php/tests/php/test_php_web_htmlentities_charset_quotes.rs

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

$str = "<p> 'single' & \"double\" </p>";
echo htmlentities($str, ENT_QUOTES, 'UTF-8'), "\n";

__vybe_check(ob_get_clean(), "&lt;p&gt; &#039;single&#039; &amp; &quot;double&quot; &lt;/p&gt;");
