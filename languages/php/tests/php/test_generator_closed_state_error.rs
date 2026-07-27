crate::php_cases! {
    generator_closed_state_error => {
        r#"<?php
function gen() {
    yield 1;
}
$g = gen();
$g->next();
// generator is now closed
try {
    $g->next();
    echo "ok";
} catch (\Exception $e) {
    echo "error";
} catch (\Error $e) {
    echo "error";
}
"#,
        ["ok"] // Calling next() on closed generator does nothing or throws? Actually it does nothing, returns null.
    };

    generator_getreturn_on_closed => {
        r#"<?php
function gen() { yield 1; }
$g = gen();
$g->next();
echo is_null($g->getReturn()) ? "null" : "not";
"#,
        ["null"]
    };
}
