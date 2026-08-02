// vybe-test: kotlin/local_returns/test_local_return_in_repeat_with_indices
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            repeat(3) {
                if (it == 1) return@repeat
                out.append(it)
            }
            __check((out.toString()).toString(), "02")
        }
