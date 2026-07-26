
crate::php_cases! {
    generator_rewind_after_start_throws => {
        r#"<?php
function gen() { yield 1; yield 2; }
$g = gen();
$g->next();
try {
    $g->rewind();
    echo "ok";
} catch (\Exception $e) {
    echo "error";
}
"#,
        ["error"]
    };

    generator_rewind_before_start => {
        r#"<?php
function gen() { yield 1; }
$g = gen();
$g->rewind();
echo $g->current();
"#,
        ["1"]
    };
}
