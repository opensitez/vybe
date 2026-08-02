// vybe-test: kotlin/functions/test_function_multiple_parameters
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun add(a: Int, b: Int, c: Int): Int {
            return a + b + c
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((add(1, 2, 3)).toString(), "6")
        }
