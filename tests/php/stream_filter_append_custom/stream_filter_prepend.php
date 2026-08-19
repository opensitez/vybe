<?php
// vybe-test: php/stream_filter_append_custom/stream_filter_prepend
// origin: languages/php/tests/php/test_stream_filter_append_custom.rs

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

class PrefixFilter extends php_user_filter {
    public function filter($in, $out, &$consumed, $closing) {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = "PRE_" . $bucket->data;
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }
        return PSFS_PASS_ON;
    }
}
stream_filter_register("string.prefix", "PrefixFilter");

$fp = fopen('php://memory', 'w+');
stream_filter_prepend($fp, "string.prefix");
fwrite($fp, "data");
rewind($fp);
echo stream_get_contents($fp);
fclose($fp);

__vybe_check(ob_get_clean(), "PRE_PRE_data");
