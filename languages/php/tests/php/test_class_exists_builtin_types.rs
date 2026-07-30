//! `class_exists` / `interface_exists` on types the compiler does NOT declare.
//!
//! `class_exists('PDO')` is the standard PHP feature-detection idiom, so it has
//! to stay true for host- and prelude-provided types, not just for classes the
//! user wrote. These probe the runtime `__kind` annotation, and only a type that
//! went through the shared class compiler carries one — so an absent stamp must
//! fall back to "is it defined", never answer false.

crate::php_cases! {
    class_exists_on_builtin_exception => {
        r#"<?php
echo class_exists('Exception') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_exists_on_stdclass => {
        r#"<?php
echo class_exists('stdClass') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_exists_on_user_class_still_true => {
        r#"<?php
class PlainUserClass {}
echo class_exists('PlainUserClass') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    // The discrimination that the kind annotation exists for must survive the
    // absent-stamp fallback: a declared interface is NOT a class.
    interface_is_not_a_class => {
        r#"<?php
interface DeclaredContract {}
echo class_exists('DeclaredContract') ? 'cls' : 'notcls';
echo '|';
echo interface_exists('DeclaredContract') ? 'iface' : 'noiface';
"#,
        ["notcls|iface"]
    };

    // And a genuinely missing name stays false whichever way it is asked.
    missing_type_is_false_for_every_kind => {
        r#"<?php
echo class_exists('NoSuchTypeAnywhere', false) ? 'y' : 'n';
echo interface_exists('NoSuchTypeAnywhere', false) ? 'y' : 'n';
echo trait_exists('NoSuchTypeAnywhere', false) ? 'y' : 'n';
"#,
        ["nnn"]
    };
}
