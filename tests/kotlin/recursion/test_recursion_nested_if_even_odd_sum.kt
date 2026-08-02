// vybe-test: kotlin/recursion/test_recursion_nested_if_even_odd_sum
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun altSum(n: Int): Int = if (n <= 0) 0 else if (n % 2 == 0) altSum(n - 1) - n else altSum(n - 1) + n
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((altSum(4)).toString(), "2")
        }
