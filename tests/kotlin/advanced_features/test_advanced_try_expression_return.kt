// vybe-test: kotlin/advanced_features/test_advanced_try_expression_return
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun compute(): Int {
            return try {
                20 / 2
            } catch (e: Exception) {
                0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute()).toString(), "10")
        }
