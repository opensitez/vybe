crate::php_cases! {
    generator_nested_yield_from => {
        r#"<?php
function gen3() { yield 3; }
function gen2() { yield 2; yield from gen3(); }
function gen1() { yield 1; yield from gen2(); yield 4; }

$out = [];
foreach (gen1() as $v) $out[] = $v;
echo implode(',', $out);
"#,
        ["1,2,3,4"]
    };
}
