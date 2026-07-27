crate::php_cases! {
    generator_memory_cleanup => {
        r#"<?php
class ResourceObj {
    public static $count = 0;
    public function __construct() { self::$count++; }
    public function __destruct() { self::$count--; }
}

function gen() {
    $obj = new ResourceObj();
    yield 1;
    yield 2;
}
$g = gen();
$g->current();
echo ResourceObj::$count . "|";
$g = null;
echo ResourceObj::$count;
"#,
        ["1|0"]
    };
}
