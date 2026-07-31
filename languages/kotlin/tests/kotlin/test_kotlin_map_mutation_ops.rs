use crate::helpers::run_prints;

#[test]
fn test_mutable_map_put_and_update() {
    let out = run_prints(r#"
        fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = 2
            m.put("b", 3)
            println(m["a"])
            println(m["b"])
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_mutable_map_removal() {
    let out = run_prints(r#"
        fun main() {
            val m = mutableMapOf("x" to 1, "y" to 2)
            m.remove("x")
            println(m.containsKey("x"))
            println(m.size)
        }
    "#);
    assert_eq!(out, &["false", "1"]);
}
