// vybe-test: kotlin/builtins/test_math_coerce_in_invalid_bounds_throws
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun main() {
            try {
                println(1.coerceIn(5, 2))
            } catch (e: IllegalArgumentException) {
                println("invalid")
            }
        }

