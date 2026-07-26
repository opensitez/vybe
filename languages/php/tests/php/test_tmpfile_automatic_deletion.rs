
crate::php_cases! {
    tmpfile_creation => {
        r#"<?php
$temp = tmpfile();
fwrite($temp, "test data");
rewind($temp);
echo fread($temp, 1024);
fclose($temp); // should delete the file
"#,
        ["test data"]
    };
}
