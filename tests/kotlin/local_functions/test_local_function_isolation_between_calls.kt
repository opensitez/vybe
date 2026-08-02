// vybe-test: kotlin/local_functions/test_local_function_isolation_between_calls
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun runOnce(v: Int): Int {
                fun bump(x: Int): Int = x + 1
                return bump(v)
            }
            __check((runOnce(1)).toString(), "2")
            __check((runOnce(2)).toString(), "3")
        }
