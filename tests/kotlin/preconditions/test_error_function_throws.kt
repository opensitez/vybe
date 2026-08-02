// vybe-test: kotlin/preconditions/test_error_function_throws
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                error("fatal")
                println("no")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }

