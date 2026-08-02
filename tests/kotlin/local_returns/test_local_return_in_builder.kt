// vybe-test: kotlin/local_returns/test_local_return_in_builder
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = StringBuilder().apply {
                append("a")
                if (length == 0) return
                append("b")
            }
            __check((result.toString()).toString(), "ab")
        }
