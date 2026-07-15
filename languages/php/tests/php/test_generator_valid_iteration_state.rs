use super::helpers::run_prints;

crate::php_cases! {
    generator_valid_iteration => {
        r#"<?php
function gen() {
    yield 1;
}
$g = gen();
echo ($g->valid() ? 'yes' : 'no') . "|";
$g->next();
echo ($g->valid() ? 'yes' : 'no');
"#,
        ["yes|no"]
    };
}
