kotlin_run_cases! {
    test_infix_to_pair => (r#"
        fun main() {
            val pair = 1 to 2
            println(pair.first)
            println(pair.second)
        }
    "#, vec!["1", "2"]),
    test_infix_is_operator => (r#"
        fun main() {
            val x: Any = "x"
            val y = "x"
            println((x is String).toString())
            println((x !is Int).toString())
            println((y is String).toString())
        }
    "#, vec!["true", "true", "true"]),
    test_infix_else_chain => (r#"
        fun classify(v: Int): String {
            return if (v % 2 == 0 && v > 0) {
                "even"
            } else {
                "odd"
            }
        }

        fun main() {
            println(classify(4))
            println(classify(5))
        }
    "#, vec!["even", "odd"]),
}
