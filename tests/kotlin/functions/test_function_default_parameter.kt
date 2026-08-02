// vybe-test: kotlin/functions/test_function_default_parameter
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun greet(name: String = "friend") {
            println("hi " + name)
        }

        fun main() {
            greet()
            greet("Dev")
        }

