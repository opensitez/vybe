// vybe-test: kotlin/recursion/test_recursion_sum_down
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun sum(n: Int): Int = if (n <= 0) 0 else n + sum(n - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sum(4)).toString(), "10")
        }
