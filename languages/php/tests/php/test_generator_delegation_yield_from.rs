crate::php_cases! {
    generator_yield_from_array => {
        r#"<?php
function gen() {
    yield from [1, 2];
    yield 3;
}
$out = [];
foreach (gen() as $v) {
    $out[] = $v;
}
echo implode(',', $out);
"#,
        ["1,2,3"]
    };

    generator_yield_from_generator => {
        r#"<?php
function inner() { yield 'a'; yield 'b'; return 'c'; }
function outer() {
    $ret = yield from inner();
    yield $ret;
}
$out = [];
foreach (outer() as $v) $out[] = $v;
echo implode(',', $out);
"#,
        ["a,b,c"]
    };
}
