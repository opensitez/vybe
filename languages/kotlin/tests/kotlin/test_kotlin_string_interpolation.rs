kotlin_run_cases! {
    test_simple_interpolation => (r#"
        fun main() {
            val name = "kotlin"
            println("hi $name")
        }
    "#, vec!["hi kotlin"]),
    test_expression_interpolation => (r#"
        fun main() {
            val value = 3
            println("sum " + "${value + 1}")
        }
    "#, vec!["sum 4"]),
    test_property_interpolation => (r#"
        class User(val name: String, val age: Int)

        fun main() {
            val u = User("bob", 5)
            println("${u.name}:${u.age}")
        }
    "#, vec!["bob:5"]),
}
