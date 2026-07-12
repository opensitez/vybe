//! Magic constants `__LINE__`, `__FILE__`, `__DIR__`, `__FUNCTION__`, `__METHOD__`.

crate::php_cases! {
    line_magic_constant_is_positive_integer => {
        r#"<?php
echo __LINE__ > 0 ? 'line' : 'zero';
"#,
        ["line"]
    };

    file_magic_constant_non_empty_string => {
        r#"<?php
echo strlen(__FILE__) > 0 ? 'file' : 'empty';
"#,
        ["file"]
    };

    dir_magic_constant_non_empty_string => {
        r#"<?php
echo strlen(__DIR__) > 0 ? 'dir' : 'empty';
"#,
        ["dir"]
    };

    function_magic_inside_named_function => {
        r#"<?php
function named(): string { return __FUNCTION__; }
echo named();
"#,
        ["named"]
    };

    method_magic_inside_class_method => {
        r#"<?php
class M { public function go(): string { return __METHOD__; } }
echo (new M())->go();
"#,
        ["M::go"]
    };

    class_magic_inside_class => {
        r#"<?php
class C { public function name(): string { return __CLASS__; } }
echo (new C())->name();
"#,
        ["C"]
    };

    trait_magic_inside_trait_method => {
        r#"<?php
trait T { public function m(): string { return __TRAIT__; } }
class U { use T; }
echo (new U())->m();
"#,
        ["T"]
    };

    namespace_magic_inside_namespace => {
        r#"<?php
namespace App\Magic {
    function ns(): string { return __NAMESPACE__; }
}
echo \App\Magic\ns();
"#,
        ["App\\Magic"]
    };
}
