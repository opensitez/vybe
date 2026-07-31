use crate::helpers::run_prints;

#[test]
fn test_map_keys_and_values_views() {
    let out = run_prints(r#"
        fun main() {
            val m = mapOf("x" to 1, "y" to 2)
            println(m.keys.toString())
            println(m.values.toString())
        }
    "#);
    assert_eq!(out, &["[x, y]", "[1, 2]"]);
}

#[test]
fn test_map_filtering_projection() {
    let out = run_prints(r#"
        fun main() {
            val m = mapOf("a" to 1, "b" to 3, "c" to 5)
            val f = m.filterValues { it > 1 }
            println(f["b"])
            println(f["a"])
        }
    "#);
    assert_eq!(out, &["3", "null"]);
}
