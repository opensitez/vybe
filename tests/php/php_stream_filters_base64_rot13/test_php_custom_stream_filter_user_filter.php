<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_custom_stream_filter_user_filter
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs
// vybe-test-mode: compile

class StripVowelsFilter extends php_user_filter {
    public function filter($in, $out, &$consumed, $closing): int {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = preg_replace('/[aeiouAEIOU]/', '', $bucket->data);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }
        return PSFS_PASS_ON;
    }
}

stream_filter_register("strip_vowels", StripVowelsFilter::class);
$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "strip_vowels");

fwrite($stream, "Hello World");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
