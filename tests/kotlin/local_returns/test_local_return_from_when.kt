// vybe-test: kotlin/local_returns/test_local_return_from_when
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun status(v: Int): String {
            return when (v) {
                in 1..3 -> "small"
                else -> return "other"
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((status(2)).toString(), "small")
            __check((status(10)).toString(), "other")
        }
