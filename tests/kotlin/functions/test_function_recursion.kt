// vybe-test: kotlin/functions/test_function_recursion
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun fact(n: Int): Int {
            if (n <= 1) {
                return 1
            }
            return n * fact(n - 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((fact(5)).toString(), "120")
        }
