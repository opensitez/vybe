// vybe-test: kotlin/exceptions/test_check_false_path_in_nested_context
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun assertPositive(n: Int) {
            check(n > 0)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                assertPositive(0)
            } catch (e: Exception) {
                __check(("invalid").toString(), "invalid")
            }
        }
