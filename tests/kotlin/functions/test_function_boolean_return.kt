// vybe-test: kotlin/functions/test_function_boolean_return
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun isEven(n: Int): Boolean = (n % 2 == 0)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isEven(4)).toString(), "true")
            __check((isEven(7)).toString(), "false")
        }
