kotlin_run_cases! {
    test_arithmetic_precedence => (r#"
        fun main() {
            println(1 + 2 * 3)
            println((1 + 2) * 3)
            println(10 - 2 - 1)
            println(10 - (2 - 1))
        }
    "#, vec!["7", "9", "7", "9"]),
    test_boolean_precedence => (r#"
        fun main() {
            println(true || false && false)
            println((true || false) && false)
        }
    "#, vec!["true", "false"]),
    test_ternary_like_precedence => (r#"
        fun main() {
            val x = 5
            val y = 2
            println(x + y * 2 - 1)
            println((x + y) * (2 - 1))
        }
    "#, vec!["8", "7"]),
}
