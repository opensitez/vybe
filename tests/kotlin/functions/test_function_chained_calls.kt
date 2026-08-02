// vybe-test: kotlin/functions/test_function_chained_calls
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun inc(x: Int): Int = x + 1

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((inc(inc(inc(0)))).toString(), "3")
        }
