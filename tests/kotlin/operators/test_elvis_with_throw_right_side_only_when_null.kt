// vybe-test: kotlin/operators/test_elvis_with_throw_right_side_only_when_null
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun fail(reason: String): Nothing {
            throw Exception(reason)
        }

        fun main() {
            val value: String? = "ok"
            println(value ?: fail("oops"))
            val missing: String? = null
            try {
                println(missing ?: fail("missing"))
            } catch (e: Exception) {
                println("caught")
            }
        }

