use crate::helpers::run_prints;

#[test]
fn test_nested_named_object_is_accessible() {
    let out = run_prints(
        r#"
        object Container {
            object Tag {
                fun value(): String = "nested"
            }
        }

        fun main() {
            println(Container.Tag.value())
        }
    "#,
    );
    assert_eq!(out, &["nested"]);
}

#[test]
fn test_anonymous_object_captures_outer_scope() {
    let out = run_prints(
        r#"
        fun main() {
            val label = "a"
            val obj = object {
                fun render(): String = label + "b"
            }
            println(obj.render())
        }
    "#,
    );
    assert_eq!(out, &["ab"]);
}
