use super::helpers::run_prints;

crate::php_cases! {
    generator_get_return_value => {
        r#"<?php
function gen() {
    yield 1;
    return "done";
}
$g = gen();
foreach ($g as $val) {}
echo $g->getReturn();
"#,
        ["done"]
    };

    generator_get_return_error_if_not_finished => {
        r#"<?php
function gen() {
    yield 1;
    return "done";
}
$g = gen();
try {
    echo $g->getReturn();
} catch (\Exception $e) {
    echo "error";
} catch (\Error $e) {
    echo "error";
}
"#,
        ["error"]
    };
}
