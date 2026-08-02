// vybe-test: kotlin/preconditions/test_error_before_return_is_caught
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun risky(v: Int): Int {
            require(v > 0)
            return v
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = try {
                risky(0)
            } catch (e: IllegalArgumentException) {
                -1
            }
            __check((value).toString(), "-1")
        }
