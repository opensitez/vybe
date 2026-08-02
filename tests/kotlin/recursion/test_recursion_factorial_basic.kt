// vybe-test: kotlin/recursion/test_recursion_factorial_basic
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun fact(n: Int): Int = if (n <= 1) 1 else n * fact(n - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((fact(5)).toString(), "120")
        }
