
crate::php_cases! {
    generator_current_key_sync => {
        r#"<?php
function gen() {
    yield 'a' => 1;
    yield 'b' => 2;
}
$g = gen();
echo $g->key() . ":" . $g->current() . "|";
$g->next();
echo $g->key() . ":" . $g->current();
"#,
        ["a:1|b:2"]
    };
}
