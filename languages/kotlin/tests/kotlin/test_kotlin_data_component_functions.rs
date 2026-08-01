use crate::helpers::run_prints;

#[test]
fn test_data_class_components_are_indexable() {
    let out = run_prints(
        r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val point = Point(4, 7)
            println(point.component1())
            println(point.component2())
        }
    "#,
    );
    assert_eq!(out, &["4", "7"]);
}

#[test]
fn test_data_class_copy_preserves_and_changes_fields() {
    let out = run_prints(
        r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val point = Point(2, 3)
            val shifted = point.copy(y = 9)
            println(point)
            println(shifted)
        }
    "#,
    );
    assert_eq!(out, &["Point(x=2, y=3)", "Point(x=2, y=9)"]);
}
