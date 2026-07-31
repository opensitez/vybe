kotlin_run_cases! {
    test_boolean_operators => (r#"
        fun main() {
            println(true)
            println(false)
            println(true && false)
            println(true || false)
            println(!true)
        }
    "#, vec!["true", "false", "false", "true", "false"]),
    test_boolean_comparisons => (r#"
        fun main() {
            println(1 == 1)
            println(1 != 2)
            println(2 > 1)
            println(3 <= 3)
        }
    "#, vec!["true", "true", "true", "true"]),
    test_boolean_condition_branching => (r#"
        fun main() {
            val x = if (true) {
                1
            } else {
                2
            }
            println(x)
        }
    "#, vec!["1"]),
}
