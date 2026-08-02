// vybe-test: kotlin/operators/test_not_null_assertion_throws_on_null_reference
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun main() {
            val missing: String? = null
            try {
                println(missing!!)
            } catch (e: NullPointerException) {
                println("null")
            }
        }

