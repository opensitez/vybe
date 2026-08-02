// vybe-test: kotlin/local_returns/test_local_return_in_repeat
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            repeat(4) { index ->
                if (index == 2) return@repeat
                out.append(index)
            }
            __check((out.toString()).toString(), "013")
        }
