// vybe-test: kotlin/nullability/test_null_assertion_failed
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun main() {
            val user: String? = null
            try {
                println(user!!)
            } catch (e: Exception) {
                println("null")
            }
        }

