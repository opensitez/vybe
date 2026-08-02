// vybe-test: kotlin/functions/test_function_local_nested
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun triple(x: Int): Int {
                return x * 3
            }
            __check((triple(4)).toString(), "12")
        }
