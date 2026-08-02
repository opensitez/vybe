// vybe-test: kotlin/recursion/test_recursion_mutual_even_odd
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun even(n: Int): Boolean = if (n == 0) true else odd(n - 1)
        fun odd(n: Int): Boolean = if (n == 0) false else even(n - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((even(10)).toString(), "true")
            __check((odd(10)).toString(), "false")
        }
