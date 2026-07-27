crate::php_cases! {
    generator_send_resume => {
        r#"<?php
function gen() {
    $in = yield 'first';
    yield $in . ' received';
}
$g = gen();
echo $g->current() . "|";
$g->send('hello');
echo $g->current();
"#,
        ["first|hello received"]
    };
}
