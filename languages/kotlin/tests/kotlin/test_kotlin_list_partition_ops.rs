use crate::helpers::run_prints;

#[test]
fn test_list_partition_by_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val src = listOf(1, 2, 3, 4)
            val p = src.partition { it % 2 == 0 }
            println(p.first.toString())
            println(p.second.toString())
        }
    "#,
    );
    assert_eq!(out, &["[2, 4]", "[1, 3]"]);
}

#[test]
fn test_list_drop_take_bounds() {
    let out = run_prints(
        r#"
        fun main() {
            val src = listOf("a", "b", "c")
            println(src.drop(1).toString())
            println(src.take(2).toString())
        }
    "#,
    );
    assert_eq!(out, &["[b, c]", "[a, b]"]);
}
