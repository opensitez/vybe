use super::helpers::run_prints;

crate::php_cases! {
    generator_throw_exception_caught => {
        r#"<?php
function gen() {
    try {
        yield 1;
    } catch (\Exception $e) {
        yield $e->getMessage();
    }
}
$g = gen();
echo $g->current() . "|";
$g->throw(new \Exception("thrown"));
echo $g->current();
"#,
        ["1|thrown"]
    };
}
