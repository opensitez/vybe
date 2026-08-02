// vybe-test: kotlin/functions/test_function_default_and_named_args
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun greet(prefix: String = "Hi", name: String) {
            println(prefix + " " + name)
        }

        fun main() {
            greet(name = "Kotlin")
            greet(prefix = "Hello", name = "Rust")
        }

