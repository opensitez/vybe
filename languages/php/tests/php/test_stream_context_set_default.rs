use super::helpers::run_prints;

crate::php_cases! {
    stream_context_set_default_opts => {
        r#"<?php
$opts = ['http' => ['method' => 'POST']];
$ctx = stream_context_set_default($opts);
$params = stream_context_get_options(stream_context_get_default());
echo $params['http']['method'];
"#,
        ["POST"]
    };
}
