use crate::helpers::run_prints;

#[test]
fn test_int_range_until_excludes_end() {
    let out = run_prints(r#"
        fun main() {
            var acc = 0
            for (i in 1 until 4) {
                acc += i
            }
            println(acc)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_int_range_step_expressions() {
    let out = run_prints(r#"
        fun main() {
            var acc = ""
            for (i in 0..6 step 2) {
                acc += i.toString()
            }
            println(acc)
        }
    "#);
    assert_eq!(out, &["0246"]);
}
