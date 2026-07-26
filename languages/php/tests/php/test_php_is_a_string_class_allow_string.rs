use super::helpers::run_prints;

#[test]
fn test_is_a_string_class_name_allow_string_true() {
    assert_eq!(
        run_prints(
            r#"<?php
class ParentType {}
class ChildType extends ParentType {}
echo is_a('ChildType', 'ParentType', true) ? 'subclass_string_ok' : 'err', "\n";
"#
        ),
        vec!["subclass_string_ok"]
    );
}

#[test]
fn test_is_a_string_class_name_allow_string_false() {
    assert_eq!(
        run_prints(
            r#"<?php
class ParentType2 {}
class ChildType2 extends ParentType2 {}
echo is_a('ChildType2', 'ParentType2', false) ? 'unexpected' : 'string_disallowed', "\n";
"#
        ),
        vec!["string_disallowed"]
    );
}
