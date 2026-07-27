crate::php_cases! {
    stream_filter_register_and_append => {
        r#"<?php
class UpperFilter extends php_user_filter {
    public function filter($in, $out, &$consumed, $closing) {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = strtoupper($bucket->data);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }
        return PSFS_PASS_ON;
    }
}
stream_filter_register("string.toupper_custom", "UpperFilter");

$fp = fopen('php://memory', 'w+');
stream_filter_append($fp, "string.toupper_custom");
fwrite($fp, "hello");
rewind($fp);
echo stream_get_contents($fp);
fclose($fp);
"#,
        ["HELLO"]
    };

    stream_filter_prepend => {
        r#"<?php
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
"#,
        ["PRE_data"]
    };

    stream_filter_remove => {
        r#"<?php
class DummyFilter extends php_user_filter {
    public function filter($in, $out, &$consumed, $closing) {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = str_replace('a', 'b', $bucket->data);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }
        return PSFS_PASS_ON;
    }
}
stream_filter_register("dummy", "DummyFilter");

$fp = fopen('php://memory', 'w+');
$filter = stream_filter_append($fp, "dummy");
fwrite($fp, "aaa");
stream_filter_remove($filter);
fwrite($fp, "ccc");
rewind($fp);
echo stream_get_contents($fp);
fclose($fp);
"#,
        ["bbbccc"]
    };
}
