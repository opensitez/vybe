crate::php_cases! {
    sscanf_basic_return_array => {
        r#"<?php
$str = "October 24, 1990";
$format = "%s %d, %d";
$res = sscanf($str, $format);
echo count($res) . "|" . $res[0] . "|" . $res[1] . "|" . $res[2];
"#,
        ["3|October|24|1990"]
    };

    sscanf_pass_by_reference => {
        r#"<?php
$str = "Author: John Doe";
$format = "Author: %s %s";
$count = sscanf($str, $format, $first, $last);
echo $count . "|" . $first . "|" . $last;
"#,
        ["2|John|Doe"]
    };
}
