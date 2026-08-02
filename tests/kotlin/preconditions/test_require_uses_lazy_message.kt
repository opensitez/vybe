// vybe-test: kotlin/preconditions/test_require_uses_lazy_message
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val side = arrayOf(0)
            try {
                require(side[0] > 0, { side[0] = 1; "message" })
            } catch (e: IllegalArgumentException) {
                __check((side[0]).toString(), "1")
                __check((e.message).toString(), "message")
            }
        }
