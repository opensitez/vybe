// vybe-test: kotlin/nullability/test_null_check_if_else
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun main() {
            val str: String? = null
            if (str != null) {
                println("valid")
            } else {
                println("null value")
            }
        }

