// vybe-test: kotlin/local_functions/test_local_function_reassignment_not_allowed_and_compiler_checks
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun main() {
            fun square(x: Int): Int = x * x
            try {
                println(square(4))
            } catch (error: Exception) {
                println("bad")
            }
        }

