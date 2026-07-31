use crate::helpers::run_prints;

#[test]
fn test_array_indexing_and_size() {
    let out = run_prints(r#"
        fun main() {
            val a = arrayOf(10, 20, 30)
            println(a.size)
            println(a[1] + a[2])
        }
    "#);
    assert_eq!(out, &["3", "50"]);
}

#[test]
fn test_array_iteration_sum() {
    let out = run_prints(r#"
        fun main() {
            val a = arrayOf(2, 4, 6)
            var acc = 0
            for (v in a) {
                acc += v
            }
            println(acc)
        }
    "#);
    assert_eq!(out, &["12"]);
}
