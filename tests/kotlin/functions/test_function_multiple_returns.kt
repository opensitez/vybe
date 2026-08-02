// vybe-test: kotlin/functions/test_function_multiple_returns
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun sign(x: Int): Int {
            if (x > 0) return 1
            if (x < 0) return -1
            return 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sign(10)).toString(), "1")
            __check((sign(-5)).toString(), "-1")
            __check((sign(0)).toString(), "0")
        }
