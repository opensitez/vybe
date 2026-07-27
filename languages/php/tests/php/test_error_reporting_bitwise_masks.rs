crate::php_cases! {
    error_reporting_bitwise => {
        r#"<?php
$old = error_reporting(E_ALL & ~E_NOTICE & ~E_USER_NOTICE);

echo error_reporting() === (E_ALL & ~E_NOTICE & ~E_USER_NOTICE) ? "ok|" : "fail|";

error_reporting($old);
echo "restored";
"#,
        ["ok|restored"]
    };
}
