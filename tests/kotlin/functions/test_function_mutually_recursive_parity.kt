// vybe-test: kotlin/functions/test_function_mutually_recursive_parity
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun isEven(n: Int): Boolean {
            if (n == 0) return true
            return isOdd(n - 1)
        }

        fun isOdd(n: Int): Boolean {
            if (n == 0) return false
            return isEven(n - 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isEven(8)).toString(), "true")
            __check((isOdd(8)).toString(), "false")
            __check((isEven(7)).toString(), "false")
        }
