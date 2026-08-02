// vybe-test: kotlin/local_functions/test_local_function_recursive_factorial
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun fact(n: Int): Int {
                return if (n <= 1) 1 else n * fact(n - 1)
            }
            __check((fact(5)).toString(), "120")
        }
