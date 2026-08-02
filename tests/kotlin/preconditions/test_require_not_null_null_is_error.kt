// vybe-test: kotlin/preconditions/test_require_not_null_null_is_error
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                val value: String? = null
                requireNotNull(value)
                println("no")
            } catch (e: IllegalArgumentException) {
                println("missing")
            }
        }

