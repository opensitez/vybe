// vybe-test: kotlin/functions/test_function_boolean_contract
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun allPositive(a: Int, b: Int): Boolean {
            return a > 0 && b > 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((allPositive(3, 4)).toString(), "true")
            __check((allPositive(3, -1)).toString(), "false")
        }
