crate::php_cases! {
    fscanf_basic_read => {
        r#"<?php
$fp = fopen("php://memory", "w+");
fwrite($fp, "101 John\n102 Jane");
rewind($fp);

$user1 = fscanf($fp, "%d %s");
$user2 = fscanf($fp, "%d %s");
echo $user1[0] . "-" . $user1[1] . "|" . $user2[0] . "-" . $user2[1];
fclose($fp);
"#,
        ["101-John|102-Jane"]
    };

    fscanf_pass_by_reference => {
        r#"<?php
$fp = fopen("php://memory", "w+");
fwrite($fp, "Color: Red\n");
rewind($fp);

$count = fscanf($fp, "Color: %s", $color);
echo $count . "|" . $color;
fclose($fp);
"#,
        ["1|Red"]
    };
}
