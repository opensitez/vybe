// vybe-test: kotlin/numeric_types/test_integer_division_by_zero_is_runtime_error
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun main() {
            try {
                println(7 / 0)
            } catch (e: Exception) {
                println("division-error")
            }
        }

