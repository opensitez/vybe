use crate::helpers::run_prints;

#[test]
fn test_long_array_literals_and_arith() {
    let out = run_prints(
        r#"
        fun main() {
            val a = longArrayOf(1L, 2L, 3L)
            println(a[2] - a[0])
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_long_array_size_growth_like_mutation() {
    let out = run_prints(
        r#"
        fun main() {
            val a = longArrayOf(9L, 8L)
            val b = a.copyOf(3)
            b[2] = 7L
            println(b.size)
            println(b[2])
        }
    "#,
    );
    assert_eq!(out, &["3", "7"]);
}
